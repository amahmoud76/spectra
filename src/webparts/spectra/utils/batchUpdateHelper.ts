// ─────────────────────────────────────────────────────────────────
// Batch Update Metadata — pure planning & conflict-resolution logic
//
// Scope: CURRENT documents only. Archived documents are never touched.
// SET (replace) semantics: only fields the admin explicitly changes are
// applied to every matched document.
//
// Duplicate resolution (approved rule): when a change would make two
// Current documents share the same identity, keep the most recent
// effective date and archive the other. To avoid a filename collision,
// we KEEP the document that already holds the canonical name in place,
// carry the winning effective date onto it, and archive the other AS-IS
// (it keeps its original name — a Status change is sufficient).
// ─────────────────────────────────────────────────────────────────

import { IDocument } from "../interfaces/IDocument";
import { IUploadPayload } from "../interfaces/IUploadPayload";

// Metadata array fields that a batch run may change.
export type BatchMetadataField =
  | "therapeuticArea"
  | "subTherapeuticArea"
  | "diseaseArea"
  | "asset"
  | "indication"
  | "lineOfTherapy"
  | "paid";

export const BATCH_METADATA_FIELDS: BatchMetadataField[] = [
  "therapeuticArea",
  "subTherapeuticArea",
  "diseaseArea",
  "asset",
  "indication",
  "lineOfTherapy",
  "paid",
];

// The set of changes the admin wants applied to every matched document.
// Only keys present here are considered "changed" and will be written.
export interface IBatchChangeSet {
  metadata: Partial<Record<BatchMetadataField, string[]>>;
  // ISO date string when the admin chose to change the effective date.
  effectiveDate?: string;
}

export type BatchItemAction =
  // Apply the change in place (MERGE + conditional rename). No conflict.
  | "update"
  // Keep this document (already holds the canonical name); only carry over
  // the winning effective date. Used for the surviving side of a conflict.
  | "keep-with-date"
  // Archive this document as-is (it loses a duplicate conflict).
  | "archive-loser"
  // Nothing to do (change produced no difference for this document).
  | "no-op";

export interface IBatchItemPlan {
  doc: IDocument;
  action: BatchItemAction;
  // Fields to write for "update" / effective date for "keep-with-date".
  targetPayload?: Partial<IUploadPayload>;
  // The document this one collides with (for reporting).
  conflictWithId?: string;
  conflictWithFileName?: string;
}

export interface IBatchPlan {
  items: IBatchItemPlan[];
  toUpdateCount: number; // documents that will be modified (update + keep-with-date)
  toArchiveCount: number; // documents that will be archived as duplicates
  matchedCount: number; // total documents matched by the criteria
  hasConflicts: boolean;
}

// Outcome of executing a batch plan.
export interface IBatchFailure {
  id: string;
  fileName: string;
  reason: string;
}

export interface IBatchResult {
  updated: number;
  archived: number;
  failed: number;
  failures: IBatchFailure[];
}

// ─────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────

const sortedJoin = (values: string[] | undefined): string =>
  [...(values || [])]
    .map((v) => v.trim())
    .filter(Boolean)
    .sort()
    .join(";");

/**
 * Compute a document's identity key. Mirrors DocumentService identity logic:
 * excludes Effective Date, Comments, and PAID.
 */
export const getIdentityKey = (doc: {
  documentType: string;
  therapeuticArea: string[];
  subTherapeuticArea: string[];
  indication: string[];
  lineOfTherapy: string[];
  asset: string[];
  diseaseArea: string[];
}): string => {
  const docType = doc.documentType || "";
  const ta = sortedJoin(doc.therapeuticArea);
  const subTA = sortedJoin(doc.subTherapeuticArea);

  if (docType === "DAS") {
    const diseaseArea = sortedJoin(doc.diseaseArea);
    return `${docType}|${ta}|${subTA}|${diseaseArea}`;
  }

  const indication = sortedJoin(doc.indication);
  const lot = sortedJoin(doc.lineOfTherapy);
  const asset = sortedJoin(doc.asset);
  return `${docType}|${ta}|${subTA}|${indication}|${lot}|${asset}`;
};

/**
 * Apply a change set to a document, returning a new document copy with the
 * changed fields replaced. Non-changed fields are left intact.
 */
export const applyChangeSetToDoc = (
  doc: IDocument,
  changeSet: IBatchChangeSet,
): IDocument => {
  const next: IDocument = { ...doc };
  for (const field of BATCH_METADATA_FIELDS) {
    const value = changeSet.metadata[field];
    if (value !== undefined) {
      next[field] = [...value];
    }
  }
  if (changeSet.effectiveDate !== undefined) {
    next.effectiveDate = changeSet.effectiveDate;
  }
  return next;
};

/**
 * Build the partial upload payload containing only the changed fields, so
 * DocumentService.updateDocument merges just those.
 */
export const buildTargetPayload = (
  changeSet: IBatchChangeSet,
): Partial<IUploadPayload> => {
  const payload: Partial<IUploadPayload> = {};
  for (const field of BATCH_METADATA_FIELDS) {
    const value = changeSet.metadata[field];
    if (value !== undefined) {
      payload[field] = [...value];
    }
  }
  if (changeSet.effectiveDate !== undefined) {
    payload.effectiveDate = changeSet.effectiveDate;
  }
  return payload;
};

/** True when the change set contains at least one field to change. */
export const hasAnyChange = (changeSet: IBatchChangeSet): boolean =>
  Object.keys(changeSet.metadata).length > 0 ||
  changeSet.effectiveDate !== undefined;

const toTime = (isoDate: string | undefined): number => {
  if (!isoDate) return Number.NEGATIVE_INFINITY;
  const time = new Date(isoDate).getTime();
  return Number.isNaN(time) ? Number.NEGATIVE_INFINITY : time;
};

/**
 * Compare two documents by the "keep most recent" rule.
 * Returns the document that should WIN (survive as Current).
 * Tiebreak: later effective date → later upload date → the one with a date.
 */
const pickWinner = (
  a: { effectiveDate: string; uploadDate: string },
  b: { effectiveDate: string; uploadDate: string },
): "a" | "b" => {
  const effA = toTime(a.effectiveDate);
  const effB = toTime(b.effectiveDate);
  if (effA !== effB) return effA > effB ? "a" : "b";
  const upA = toTime(a.uploadDate);
  const upB = toTime(b.uploadDate);
  if (upA !== upB) return upA > upB ? "a" : "b";
  return "a";
};

// ─────────────────────────────────────────────────────────────────
// Plan builder
// ─────────────────────────────────────────────────────────────────

interface ITargetEntry {
  doc: IDocument;
  targetDoc: IDocument;
  newIdentity: string;
  changed: boolean;
}

/**
 * Build the batch plan for a set of matched CURRENT documents and a change set.
 *
 * @param matchedDocs   Current documents matched by the criteria (single doc type).
 * @param allCurrentDocs All Current documents in the library (for external conflict detection).
 * @param changeSet     The fields/values to apply.
 */
export const buildBatchPlan = (
  matchedDocs: IDocument[],
  allCurrentDocs: IDocument[],
  changeSet: IBatchChangeSet,
): IBatchPlan => {
  const matchedIds = new Set(matchedDocs.map((d) => d.id));

  // Non-matched current docs, indexed by their (unchanged) identity.
  // These are the potential external "name holders".
  const nameHolderByIdentity = new Map<string, IDocument>();
  for (const doc of allCurrentDocs) {
    if (matchedIds.has(doc.id)) continue;
    nameHolderByIdentity.set(getIdentityKey(doc), doc);
  }

  // Compute each matched doc's target (post-change) state + identity.
  const targets: ITargetEntry[] = matchedDocs.map((doc): ITargetEntry => {
    const targetDoc = applyChangeSetToDoc(doc, changeSet);
    return {
      doc,
      targetDoc,
      newIdentity: getIdentityKey(targetDoc),
      changed: getIdentityKey(doc) !== getIdentityKey(targetDoc),
    };
  });

  // Group matched docs by their new identity to find within-batch collisions.
  const byNewIdentity = new Map<string, ITargetEntry[]>();
  targets.forEach((t) => {
    const group = byNewIdentity.get(t.newIdentity) || [];
    group.push(t);
    byNewIdentity.set(t.newIdentity, group);
  });

  const items: IBatchItemPlan[] = [];
  const targetPayload = buildTargetPayload(changeSet);

  for (const [identity, group] of Array.from(byNewIdentity.entries())) {
    const externalHolder = nameHolderByIdentity.get(identity);
    const hasCollision = group.length > 1 || externalHolder !== undefined;

    if (!hasCollision) {
      // Simple case: single matched doc, no other doc claims this identity.
      const only = group[0];
      items.push({
        doc: only.doc,
        action: only.changed ? "update" : "no-op",
        targetPayload: only.changed ? targetPayload : undefined,
      });
      continue;
    }

    // ── Conflict resolution ────────────────────────────────────
    // Determine the surviving effective date (winner by recency), considering
    // all colliding candidates (matched docs use their TARGET effective date;
    // the external holder uses its own).
    const candidates: Array<{ effectiveDate: string; uploadDate: string }> =
      group.map((t) => ({
        effectiveDate: t.targetDoc.effectiveDate,
        uploadDate: t.doc.uploadDate,
      }));
    if (externalHolder) {
      candidates.push({
        effectiveDate: externalHolder.effectiveDate,
        uploadDate: externalHolder.uploadDate,
      });
    }
    let winningEffectiveDate = candidates[0];
    for (let i = 1; i < candidates.length; i++) {
      winningEffectiveDate =
        pickWinner(winningEffectiveDate, candidates[i]) === "a"
          ? winningEffectiveDate
          : candidates[i];
    }

    if (externalHolder) {
      // The external document already holds the canonical name → keep it in
      // place, carry the winning effective date, archive all matched docs as-is.
      const needsDateUpdate =
        toTime(externalHolder.effectiveDate) !==
        toTime(winningEffectiveDate.effectiveDate);
      items.push({
        doc: externalHolder,
        action: needsDateUpdate ? "keep-with-date" : "no-op",
        targetPayload: needsDateUpdate
          ? { effectiveDate: winningEffectiveDate.effectiveDate }
          : undefined,
        conflictWithId: group[0].doc.id,
        conflictWithFileName: group[0].doc.fileName,
      });
      for (const t of group) {
        items.push({
          doc: t.doc,
          action: "archive-loser",
          conflictWithId: externalHolder.id,
          conflictWithFileName: externalHolder.fileName,
        });
      }
    } else {
      // Within-batch collision only: the canonical name is free (no doc holds
      // it yet). Pick one matched doc to become the survivor (update + rename),
      // archive the rest as-is.
      let survivor = group[0];
      for (let i = 1; i < group.length; i++) {
        survivor =
          pickWinner(
            { effectiveDate: survivor.targetDoc.effectiveDate, uploadDate: survivor.doc.uploadDate },
            { effectiveDate: group[i].targetDoc.effectiveDate, uploadDate: group[i].doc.uploadDate },
          ) === "a"
            ? survivor
            : group[i];
      }
      for (const t of group) {
        if (t === survivor) {
          items.push({
            doc: t.doc,
            action: t.changed ? "update" : "no-op",
            targetPayload: t.changed ? targetPayload : undefined,
          });
        } else {
          items.push({
            doc: t.doc,
            action: "archive-loser",
            conflictWithId: survivor.doc.id,
            conflictWithFileName: survivor.doc.fileName,
          });
        }
      }
    }
  }

  const toUpdateCount = items.filter(
    (i) => i.action === "update" || i.action === "keep-with-date",
  ).length;
  const toArchiveCount = items.filter(
    (i) => i.action === "archive-loser",
  ).length;

  return {
    items,
    toUpdateCount,
    toArchiveCount,
    matchedCount: matchedDocs.length,
    hasConflicts: toArchiveCount > 0,
  };
};
