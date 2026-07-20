import { useState, useCallback } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDocument } from "../interfaces/IDocument";
import { IFilterState, defaultFilterState } from "../interfaces/IFilterState";
import { DocumentService } from "../services/DocumentService";
import { applyFilters } from "../utils/filterHelper";
import {
  IBatchChangeSet,
  IBatchPlan,
  IBatchResult,
  buildBatchPlan,
  hasAnyChange,
} from "../utils/batchUpdateHelper";

export type BatchStage = "criteria" | "results";

export interface IBatchProgress {
  done: number;
  total: number;
  currentFileName: string;
}

const EMPTY_CHANGE_SET: IBatchChangeSet = { metadata: {} };

export interface IUseBatchUpdateResult {
  stage: BatchStage;
  criteria: IFilterState;
  matchedDocs: IDocument[];
  changeSet: IBatchChangeSet;
  isFinding: boolean;
  isProcessing: boolean;
  progress: IBatchProgress | null;
  result: IBatchResult | null;

  setCriteria: (criteria: IFilterState) => void;
  findDocuments: () => Promise<void>;
  backToCriteria: () => void;
  setChangeSet: (changeSet: IBatchChangeSet) => void;
  /** Compute the plan for the current matched set + change set. */
  computePlan: () => IBatchPlan;
  /** Execute a plan and store the result. */
  applyPlan: (plan: IBatchPlan) => Promise<IBatchResult>;
  /** Clear everything back to a fresh Step 1 (after success or full cancel). */
  reset: () => void;
  /** Reset the working form to a fresh Step 1 but KEEP the result banner. */
  prepareNextRun: () => void;
  /** Dismiss just the success banner without clearing the (fresh) form. */
  clearResult: () => void;
}

/**
 * Orchestrates the admin Batch Update Metadata flow.
 *
 * Scope: CURRENT documents only — archived documents are never modified.
 * State lives here (mounted at the app level) so it survives the panel being
 * closed and reopened, and only clears after a successful apply or explicit reset.
 */
export const useBatchUpdate = (
  context: WebPartContext,
  documentLibrary?: string,
  useMock: boolean = false,
): IUseBatchUpdateResult => {
  const [stage, setStage] = useState<BatchStage>("criteria" as BatchStage);
  const [criteria, setCriteriaState] = useState<IFilterState>(defaultFilterState);
  const [allCurrentDocs, setAllCurrentDocs] = useState([] as IDocument[]);
  const [matchedDocs, setMatchedDocs] = useState([] as IDocument[]);
  const [changeSet, setChangeSetState] = useState<IBatchChangeSet>(EMPTY_CHANGE_SET);
  const [isFinding, setIsFinding] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState<IBatchProgress | null>(null);
  const [result, setResult] = useState<IBatchResult | null>(null);

  const setCriteria = useCallback((next: IFilterState) => {
    setCriteriaState(next);
  }, []);

  const setChangeSet = useCallback((next: IBatchChangeSet) => {
    setChangeSetState(next);
  }, []);

  const findDocuments = useCallback(async (): Promise<void> => {
    setIsFinding(true);
    try {
      const service = new DocumentService(context, documentLibrary, useMock);
      // Batch scope is Current documents only.
      const response = await service.getDocuments("Current", 1, 5000);
      const currentDocs = response.documents;
      setAllCurrentDocs(currentDocs);
      setMatchedDocs(applyFilters(currentDocs, criteria));
      setStage("results" as BatchStage);
    } catch (error) {
      console.error("useBatchUpdate.findDocuments:", error);
      setAllCurrentDocs([]);
      setMatchedDocs([]);
      setStage("results" as BatchStage);
    } finally {
      setIsFinding(false);
    }
  }, [context, documentLibrary, useMock, criteria]);

  const backToCriteria = useCallback(() => {
    setStage("criteria" as BatchStage);
  }, []);

  const computePlan = useCallback((): IBatchPlan => {
    if (!hasAnyChange(changeSet)) {
      return {
        items: [],
        toUpdateCount: 0,
        toArchiveCount: 0,
        matchedCount: matchedDocs.length,
        hasConflicts: false,
      };
    }
    return buildBatchPlan(matchedDocs, allCurrentDocs, changeSet);
  }, [matchedDocs, allCurrentDocs, changeSet]);

  const applyPlan = useCallback(
    async (plan: IBatchPlan): Promise<IBatchResult> => {
      setIsProcessing(true);
      setProgress({ done: 0, total: plan.toUpdateCount + plan.toArchiveCount, currentFileName: "" });
      try {
        const service = new DocumentService(context, documentLibrary, useMock);
        const outcome = await service.applyBatchPlan(
          plan,
          (done, total, currentFileName) =>
            setProgress({ done, total, currentFileName }),
        );
        setResult(outcome);
        return outcome;
      } finally {
        setIsProcessing(false);
        setProgress(null);
      }
    },
    [context, documentLibrary, useMock],
  );

  const reset = useCallback(() => {
    setStage("criteria" as BatchStage);
    setCriteriaState(defaultFilterState);
    setAllCurrentDocs([]);
    setMatchedDocs([]);
    setChangeSetState({ metadata: {} });
    setProgress(null);
    setResult(null);
  }, []);

  const clearResult = useCallback(() => {
    setResult(null);
  }, []);

  const prepareNextRun = useCallback(() => {
    setStage("criteria" as BatchStage);
    setCriteriaState(defaultFilterState);
    setAllCurrentDocs([]);
    setMatchedDocs([]);
    setChangeSetState({ metadata: {} });
    setProgress(null);
    // Intentionally keeps `result` so the success banner remains visible.
  }, []);

  return {
    stage,
    criteria,
    matchedDocs,
    changeSet,
    isFinding,
    isProcessing,
    progress,
    result,
    setCriteria,
    findDocuments,
    backToCriteria,
    setChangeSet,
    computePlan,
    applyPlan,
    reset,
    prepareNextRun,
    clearResult,
  };
};
