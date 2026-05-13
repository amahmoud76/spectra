import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IDocument } from "../interfaces/IDocument";
import { IDocumentResponse } from "../interfaces/IDocumentResponse";
import { IUploadPayload, IUploadResult } from "../interfaces/IUploadPayload";
import {
  USE_MOCK as DEFAULT_USE_MOCK,
  PAGE_SIZE_DEFAULT,
  SP_DOCUMENT_LIBRARY,
  MULTI_VALUE_SEPARATOR,
  SEARCH_TOKENS_SEPARATOR,
} from "../config/config";
import { mockDocuments } from "../mocks/mockDocuments";
import { generateFileName } from "../utils/fileNamingHelper";
import { formatISO, getTime } from "date-fns";

export class DocumentService {
  private context: WebPartContext;
  private _documentLibrary: string;
  private useMock: boolean;
  private mockDocs: IDocument[];

  /**
   * Compute document identity key from metadata.
   * Used to detect duplicate/replacement scenarios.
   * Format: "DocType|TA|SubTA|Indication|LOT|Asset"
   * Excludes Effective Date, Comments, PAID, and Disease Area.
   */
  private _getDocumentIdentityKey(payload: Partial<IUploadPayload>): string {
    const docType = payload.documentType || "";
    const ta = (payload.therapeuticArea || []).sort().join(";");
    const subTA = (payload.subTherapeuticArea || []).sort().join(";");
    const indication = (payload.indication || []).sort().join(";");
    const lot = (payload.lineOfTherapy || []).sort().join(";");
    const asset = (payload.asset || []).sort().join(";");
    // Note: effectiveDate, comments, paid, and diseaseArea are intentionally excluded
    return `${docType}|${ta}|${subTA}|${indication}|${lot}|${asset}`;
  }

  /**
   * Compute identity key from an existing IDocument.
   */
  private _getDocumentIdentityKeyFromDoc(doc: IDocument): string {
    const docType = doc.documentType;
    const ta = doc.therapeuticArea.sort().join(";");
    const subTA = doc.subTherapeuticArea.sort().join(";");
    const indication = doc.indication.sort().join(";");
    const lot = doc.lineOfTherapy.sort().join(";");
    const asset = doc.asset.sort().join(";");
    return `${docType}|${ta}|${subTA}|${indication}|${lot}|${asset}`;
  }

  constructor(
    context: WebPartContext,
    documentLibrary?: string,
    useMock: boolean = DEFAULT_USE_MOCK,
  ) {
    this.context = context;
    this._documentLibrary = documentLibrary || SP_DOCUMENT_LIBRARY;
    this.useMock = useMock;
    // Clone mock data so mutations don't persist across re-renders
    this.mockDocs = [...mockDocuments];
  }

  // ─────────────────────────────────────────────────────────────
  // GET DOCUMENTS
  // ─────────────────────────────────────────────────────────────

  /**
   * Fetch documents from the library.
   * @param statusFilter - 'Current' | 'Archive' | 'All'
   * @param page - 1-based page number
   * @param pageSize - items per page
   */
  public async getDocuments(
    statusFilter: "Current" | "Archive" | "All" = "All",
    page: number = 1,
    pageSize: number = PAGE_SIZE_DEFAULT,
  ): Promise<IDocumentResponse> {
    if (this.useMock) {
      await new Promise((resolve) => setTimeout(resolve, 500));

      let filtered = this.mockDocs;
      if (statusFilter !== "All") {
        filtered = filtered.filter((d) => d.status === statusFilter);
      }

      return {
        documents: filtered,
        totalCount: filtered.length,
        page,
        pageSize,
      };
    }

    // ── LIVE MODE ───────────────────────────────────────────
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items`;

    const select = [
      "Id",
      "Title",
      "FileLeafRef",
      "FileRef",
      "Created",
      "Author/Title",
      "Author/EMail",
      "Editor/Title",
      "Editor/EMail",
      "SpectraAsset",
      "SpectraDocumentType",
      "SpectraTherapeuticArea",
      "SpectraSubTherapeuticArea",
      "SpectraIndication",
      "SpectraLineOfTherapy",
      "SpectraPAID",
      "SpectraDiseaseArea",
      "SpectraEffectiveDate",
      "SpectraStatus",
      "SpectraDescription",
      "SpectraComments",
      "SpectraSearchTokens",
      "SpectraImmutableFileName",
    ].join(",");

    const expand = "Author,Editor";

    let filter = "";
    if (statusFilter !== "All") {
      filter = `&$filter=SpectraStatus eq '${statusFilter}'`;
    }

    const url = `${listUrl}?$select=${select}&$expand=${expand}&$top=5000${filter}`;

    try {
      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
        );

      if (!response.ok) {
        throw new Error(
          `DocumentService: Failed to fetch documents (${response.status})`,
        );
      }

      const data = await response.json();
      const documents: IDocument[] = data.value.map(
        (item: Record<string, unknown>) =>
          this._mapSharePointItemToDocument(item),
      );

      return {
        documents,
        totalCount: documents.length,
        page,
        pageSize,
      };
    } catch (error) {
      console.error("DocumentService.getDocuments:", error);
      throw error;
    }
  }

  // ─────────────────────────────────────────────────────────────
  // GET DOCUMENT BY ID
  // ─────────────────────────────────────────────────────────────

  public async getDocumentById(documentId: string): Promise<IDocument | null> {
    if (this.useMock) {
      await new Promise((resolve) => setTimeout(resolve, 200));
      return this.mockDocs.find((d) => d.id === documentId) || null;
    }

    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const url = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items(${documentId})?$select=*&$expand=Author,Editor`;

    try {
      const response = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
      );

      if (!response.ok) return null;

      const item = await response.json();
      return this._mapSharePointItemToDocument(item);
    } catch {
      return null;
    }
  }

  // ─────────────────────────────────────────────────────────────
  // UPLOAD DOCUMENT
  // ─────────────────────────────────────────────────────────────

  public async uploadDocument(payload: IUploadPayload, archiveTargetId?: string): Promise<IUploadResult> {
    // Generate standardized filename from metadata (BRD Appendix H)
    const generatedName = generateFileName(payload);
    // Preserve original filename for audit trail
    const originalFileName = payload.file.name;
    // Compute identity key for duplicate detection (excludes Effective Date and Comments)
    const identityKey = this._getDocumentIdentityKey(payload);

    if (this.useMock) {
      await new Promise((resolve) => setTimeout(resolve, 1000));

      // Check for existing Current document with same identity
      const duplicateCurrentDoc = this.mockDocs.find(
        (d) =>
          d.status === "Current" &&
          d.id !== archiveTargetId &&
          this._getDocumentIdentityKeyFromDoc(d) === identityKey,
      );
      if (duplicateCurrentDoc) {
        return {
          success: false,
          isDuplicateIdentity: true,
          existingDocumentId: duplicateCurrentDoc.id,
          existingFileName: duplicateCurrentDoc.fileName,
          message:
            `A current document already exists with matching metadata. ` +
            `Archive the existing document to continue, view it, or cancel.`,
        };
      }

      const archiveTarget = archiveTargetId
        ? this.mockDocs.find((d) => d.id === archiveTargetId)
        : undefined;

      const isInPlaceReplace =
        !!archiveTarget &&
        archiveTarget.fileName.trim().toLowerCase() ===
          generatedName.trim().toLowerCase();

      if (isInPlaceReplace && archiveTarget) {
        archiveTarget.fileExtension = payload.file.name.split(".").pop() || "pdf";
        archiveTarget.fileSize = payload.file.size;
        archiveTarget.modifiedBy = "Mock User";
        archiveTarget.modifiedByEmail = "mock.user@abbvie.com";
        archiveTarget.uploadDate = formatISO(new Date());
        archiveTarget.title = generatedName;
        archiveTarget.asset = payload.asset;
        archiveTarget.documentType = payload.documentType;
        archiveTarget.therapeuticArea = payload.therapeuticArea;
        archiveTarget.subTherapeuticArea = payload.subTherapeuticArea;
        archiveTarget.indication = payload.indication;
        archiveTarget.lineOfTherapy = payload.lineOfTherapy;
        archiveTarget.paid = payload.paid;
        archiveTarget.diseaseArea = payload.diseaseArea;
        archiveTarget.effectiveDate = payload.effectiveDate;
        archiveTarget.status = "Current";
        archiveTarget.description = payload.description;
        archiveTarget.comments = payload.comments;
        archiveTarget.searchTokens = payload.searchTokens;
        archiveTarget.immutableFileName = originalFileName;

        return {
          success: true,
          documentId: archiveTarget.id,
          generatedFileName: generatedName,
          message: `File successfully replaced. File name retained: ${generatedName}`,
        };
      }

      const newDoc: IDocument = {
        id: `doc-${getTime(new Date())}`,
        fileName: generatedName,
        fileExtension: payload.file.name.split(".").pop() || "pdf",
        fileUrl: `/sites/PEAKS/Documents/${generatedName}`,
        fileSize: payload.file.size,
        createdBy: "Mock User",
        createdByEmail: "mock.user@abbvie.com",
        modifiedBy: "Mock User",
        modifiedByEmail: "mock.user@abbvie.com",
        uploadDate: formatISO(new Date()),
        title: generatedName,
        asset: payload.asset,
        documentType: payload.documentType,
        therapeuticArea: payload.therapeuticArea,
        subTherapeuticArea: payload.subTherapeuticArea,
        indication: payload.indication,
        lineOfTherapy: payload.lineOfTherapy,
        paid: payload.paid,
        diseaseArea: payload.diseaseArea,
        effectiveDate: payload.effectiveDate,
        status: "Current",
        description: payload.description,
        comments: payload.comments,
        spectra511:
          this.context.pageContext.user.email || "mock.user@abbvie.com",
        searchTokens: payload.searchTokens,
        immutableFileName: originalFileName,
      };

      this.mockDocs.unshift(newDoc);

      // Archive target if provided (after upload succeeds)
      if (archiveTargetId) {
        const archiveTargetToArchive = this.mockDocs.find((d) => d.id === archiveTargetId);
        if (archiveTargetToArchive) {
          archiveTargetToArchive.status = "Archive";
        }
      }

      return {
        success: true,
        documentId: newDoc.id,
        generatedFileName: generatedName,
        message: `File successfully uploaded. File name applied: ${generatedName}`,
      };
    }

    // ── LIVE MODE ───────────────────────────────────────────
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      let replaceUsingSameGeneratedName = false;

      // Pre-flight: Check for duplicate identity before uploading file
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items`;
      const select = [
        "Id",
        "Title",
        "FileLeafRef",
        "SpectraDocumentType",
        "SpectraTherapeuticArea",
        "SpectraSubTherapeuticArea",
        "SpectraIndication",
        "SpectraLineOfTherapy",
        "SpectraAsset",
        "SpectraPAID",
        "SpectraDiseaseArea",
        "SpectraStatus",
      ].join(",");
      const dupCheckUrl = `${listUrl}?$select=${select}&$filter=SpectraStatus eq 'Current'&$top=5000`;

      const dupCheckResponse = await this.context.spHttpClient.get(
        dupCheckUrl,
        SPHttpClient.configurations.v1,
      );

      if (dupCheckResponse.ok) {
        const dupCheckData = await dupCheckResponse.json();
        const currentDocs: IDocument[] = dupCheckData.value.map(
          (item: Record<string, unknown>) =>
            this._mapSharePointItemToDocument(item),
        );

        // Check if any Current doc has same identity
        for (const currentDoc of currentDocs) {
          if (
            currentDoc.id !== archiveTargetId &&
            this._getDocumentIdentityKeyFromDoc(currentDoc) === identityKey
          ) {
            return {
              success: false,
              isDuplicateIdentity: true,
              existingDocumentId: currentDoc.id,
              existingFileName: currentDoc.fileName,
              message:
                `A current document already exists with matching metadata. ` +
                `Archive the existing document to continue, view it, or cancel.`,
            };
          }
        }
      }

      if (archiveTargetId) {
        const targetUrl = `${listUrl}(${archiveTargetId})?$select=Id,FileLeafRef,SpectraStatus`;
        const targetResponse = await this.context.spHttpClient.get(
          targetUrl,
          SPHttpClient.configurations.v1,
        );

        if (targetResponse.ok) {
          const targetData = await targetResponse.json();
          const targetFileName = String(targetData.FileLeafRef || "");
          replaceUsingSameGeneratedName =
            targetFileName.trim().toLowerCase() ===
            generatedName.trim().toLowerCase();
           // Archive the target document BEFORE uploading the replacement
           // This prevents filename conflicts in SharePoint when the new file
           // has the same name as the old one
           if (replaceUsingSameGeneratedName) {
             const archiveUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items(${archiveTargetId})`;
             await this.context.spHttpClient.post(
               archiveUrl,
               SPHttpClient.configurations.v1,
               {
                 headers: {
                   "Content-Type": "application/json;odata=nometadata",
                   "IF-MATCH": "*",
                   "X-HTTP-Method": "MERGE",
                 },
                 body: JSON.stringify({ SpectraStatus: "Archive" }),
               },
             );
           }
         }
       }

      // Step 1: Upload the file with the STANDARDIZED filename
      const fileBuffer = await payload.file.arrayBuffer();
      const uploadUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/RootFolder/Files/add(url='${encodeURIComponent(generatedName)}',overwrite=${replaceUsingSameGeneratedName ? "true" : "false"})`;

      const uploadResponse = await this.context.spHttpClient.post(
        uploadUrl,
        SPHttpClient.configurations.v1,
        {
          headers: { "Content-Type": "application/octet-stream" },
          body: fileBuffer,
        },
      );

      if (!uploadResponse.ok) {
        const uploadErrorText = await uploadResponse.text();
        const isFileNameConflict =
          uploadResponse.status === 409 ||
          /already exists|same name|name already/i.test(uploadErrorText);

        if (isFileNameConflict) {
          if (replaceUsingSameGeneratedName) {
            throw new Error(
              `Replace failed while trying to overwrite existing file: ${generatedName}. Please retry and contact support if the issue persists.`,
            );
          }

          throw new Error(
            `A document with this generated file name already exists: ${generatedName}. ` +
              "Use Replace for versioned updates, or adjust metadata fields that affect naming.",
          );
        }

        throw new Error(`Upload failed: ${uploadResponse.status}`);
      }

      const uploadData = await uploadResponse.json();

      // Step 2: Get the list item ID via the file's ServerRelativeUrl.
      // Do not URL-encode the entire server-relative path inside the OData
      // function call; SharePoint expects the raw path string there.
      const serverRelativeUrl: string = uploadData.ServerRelativeUrl;
      const escapedServerRelativeUrl = serverRelativeUrl.replace(/'/g, "''");
      const itemResponse = await this.context.spHttpClient.get(
        `${siteUrl}/_api/web/getfilebyserverrelativeurl('${escapedServerRelativeUrl}')/ListItemAllFields?$select=Id`,
        SPHttpClient.configurations.v1,
      );
      if (!itemResponse.ok) {
        throw new Error(`Uploaded file lookup failed: ${itemResponse.status}`);
      }
      const itemData = await itemResponse.json();
      const itemId = itemData.Id;

      // Step 3: Update metadata + store original filename
      const metadata = this._buildMetadataPayload(payload);
      metadata.SpectraImmutableFileName = originalFileName;
      const updateUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items(${itemId})`;

      const updateResponse = await this.context.spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Content-Type": "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify(metadata),
        },
      );
      if (!updateResponse.ok) {
        throw new Error(`Metadata update failed: ${updateResponse.status}`);
      }

      // Step 4: If archiveTargetId provided, archive that document (after upload succeeds)
      if (archiveTargetId && String(itemId) !== String(archiveTargetId)) {
         // Only archive here if filename changed (didn't archive before upload)
         if (!replaceUsingSameGeneratedName) {
           const archiveUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items(${archiveTargetId})`;
           const archiveResponse = await this.context.spHttpClient.post(
             archiveUrl,
             SPHttpClient.configurations.v1,
             {
               headers: {
                 "Content-Type": "application/json;odata=nometadata",
                 "IF-MATCH": "*",
                 "X-HTTP-Method": "MERGE",
               },
               body: JSON.stringify({ SpectraStatus: "Archive" }),
             },
           );
           if (!archiveResponse.ok) {
             throw new Error(`Archive target update failed: ${archiveResponse.status}`);
           }
         }
      }

      return {
        success: true,
        documentId: String(itemId),
        generatedFileName: generatedName,
        message: replaceUsingSameGeneratedName
          ? `File successfully replaced. File name retained: ${generatedName}`
          : `File successfully uploaded. File name applied: ${generatedName}`,
      };
    } catch (error) {
      console.error("DocumentService.uploadDocument:", error);
      return {
        success: false,
        message: `Upload failed and no file was uploaded. ${(error as Error).message}`,
      };
    }
  }

  // ─────────────────────────────────────────────────────────────
  // UPDATE DOCUMENT METADATA
  // ─────────────────────────────────────────────────────────────

  /**
   * Check if editing a document's metadata would create a duplicate identity.
   * Returns duplicate info if another Current document has the same identity with the new metadata.
   * Excludes the document being edited from the check.
   */
  public async checkEditDuplicate(
    documentId: string,
    updates: Partial<IUploadPayload>,
  ): Promise<{
    isDuplicate: boolean;
    duplicateId?: string;
    duplicateFileName?: string;
  }> {
    // Archived documents can never conflict with Current documents —
    // skip the check entirely so editing archived docs is never blocked.
    if (this.useMock) {
      const docBeingEdited = this.mockDocs.find((d) => d.id === documentId);
      if (docBeingEdited?.status === "Archive") return { isDuplicate: false };
    } else {
      try {
        const siteUrl = this.context.pageContext.web.absoluteUrl;
        const statusUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items(${documentId})?$select=SpectraStatus`;
        const statusResponse = await this.context.spHttpClient.get(
          statusUrl,
          SPHttpClient.configurations.v1,
        );
        if (statusResponse.ok) {
          const statusData = await statusResponse.json();
          if (statusData.SpectraStatus === "Archive") return { isDuplicate: false };
        }
      } catch {
        // If status fetch fails, continue with full duplicate check
      }
    }

    const buildMergedIdentityKey = (baseDoc?: IDocument): string => {
      const merged: Partial<IUploadPayload> = {
        documentType: updates.documentType ?? baseDoc?.documentType,
        therapeuticArea: updates.therapeuticArea ?? baseDoc?.therapeuticArea,
        subTherapeuticArea:
          updates.subTherapeuticArea ?? baseDoc?.subTherapeuticArea,
        indication: updates.indication ?? baseDoc?.indication,
        lineOfTherapy: updates.lineOfTherapy ?? baseDoc?.lineOfTherapy,
        asset: updates.asset ?? baseDoc?.asset,
        paid: updates.paid ?? baseDoc?.paid,
        diseaseArea: updates.diseaseArea ?? baseDoc?.diseaseArea,
      };

      return this._getDocumentIdentityKey(merged);
    };

    if (this.useMock) {
      const baseDoc = this.mockDocs.find((d) => d.id === documentId);
      const newIdentityKey = buildMergedIdentityKey(baseDoc);

      const duplicateDoc = this.mockDocs.find(
        (d) =>
          d.id !== documentId &&
          d.status === "Current" &&
          this._getDocumentIdentityKeyFromDoc(d) === newIdentityKey,
      );

      if (duplicateDoc) {
        return {
          isDuplicate: true,
          duplicateId: duplicateDoc.id,
          duplicateFileName: duplicateDoc.fileName,
        };
      }

      return { isDuplicate: false };
    }

    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items`;
      const select = [
        "Id",
        "Title",
        "FileLeafRef",
        "SpectraDocumentType",
        "SpectraTherapeuticArea",
        "SpectraSubTherapeuticArea",
        "SpectraIndication",
        "SpectraLineOfTherapy",
        "SpectraAsset",
        "SpectraPAID",
        "SpectraDiseaseArea",
        "SpectraStatus",
      ].join(",");
      // Fetch ALL documents so the base doc can be found regardless of status
      const allDocsUrl = `${listUrl}?$select=${select}&$top=5000`;

      const response = await this.context.spHttpClient.get(
        allDocsUrl,
        SPHttpClient.configurations.v1,
      );

      if (!response.ok) {
        return { isDuplicate: false };
      }

      const data = await response.json();
      const allDocs: IDocument[] = data.value.map(
        (item: Record<string, unknown>) => this._mapSharePointItemToDocument(item),
      );

      const baseDoc = allDocs.find((d) => d.id === documentId);
      const newIdentityKey = buildMergedIdentityKey(baseDoc);

      // Only flag a duplicate if another CURRENT document matches the new identity
      const duplicateDoc = allDocs.find(
        (d) =>
          d.id !== documentId &&
          d.status === "Current" &&
          this._getDocumentIdentityKeyFromDoc(d) === newIdentityKey,
      );

      if (duplicateDoc) {
        return {
          isDuplicate: true,
          duplicateId: duplicateDoc.id,
          duplicateFileName: duplicateDoc.fileName,
        };
      }

      return { isDuplicate: false };
    } catch (error) {
      console.error("DocumentService.checkEditDuplicate:", error);
      return { isDuplicate: false };
    }
  }

  public async updateDocument(
    documentId: string,
    updates: Partial<IUploadPayload>,
  ): Promise<boolean> {
    if (this.useMock) {
      await new Promise((resolve) => setTimeout(resolve, 500));
      const index = this.mockDocs.findIndex((d) => d.id === documentId);
      if (index === -1) return false;

      const doc = this.mockDocs[index];

      // Archive the old document
      doc.status = "Archive";

      // Create a new current document with updated metadata
      const newDoc: IDocument = {
        ...doc,
        id: `doc-${getTime(new Date())}`,
        status: "Current",
        uploadDate: formatISO(new Date()),
        modifiedBy: "Mock User",
        modifiedByEmail: "mock.user@abbvie.com",
        spectra511: doc.createdByEmail || doc.createdBy,
      };

      if (updates.documentType !== undefined)
        newDoc.documentType = updates.documentType;
      if (updates.therapeuticArea !== undefined)
        newDoc.therapeuticArea = updates.therapeuticArea;
      if (updates.subTherapeuticArea !== undefined)
        newDoc.subTherapeuticArea = updates.subTherapeuticArea;
      if (updates.asset !== undefined) newDoc.asset = updates.asset;
      if (updates.indication !== undefined)
        newDoc.indication = updates.indication;
      if (updates.lineOfTherapy !== undefined)
        newDoc.lineOfTherapy = updates.lineOfTherapy;
      if (updates.paid !== undefined) newDoc.paid = updates.paid;
      if (updates.diseaseArea !== undefined)
        newDoc.diseaseArea = updates.diseaseArea;
      if (updates.effectiveDate !== undefined)
        newDoc.effectiveDate = updates.effectiveDate;
      if (updates.description !== undefined)
        newDoc.description = updates.description;
      if (updates.comments !== undefined) newDoc.comments = updates.comments;
      if (updates.searchTokens !== undefined)
        newDoc.searchTokens = updates.searchTokens;

      this.mockDocs.unshift(newDoc);
      return true;
    }

    // ── LIVE MODE ───────────────────────────────────────────
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const updateUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items(${documentId})`;

      const metadata = this._buildMetadataPayload(updates);
      // Editing metadata should not change lifecycle status.
      delete metadata.SpectraStatus;

      const updateResponse = await this.context.spHttpClient.post(
        updateUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Content-Type": "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify(metadata),
        },
      );

      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        console.error("DocumentService.updateDocument failed:", {
          status: updateResponse.status,
          statusText: updateResponse.statusText,
          errorText,
        });
      }

      return updateResponse.ok;
    } catch (error) {
      console.error("DocumentService.updateDocument:", error);
      return false;
    }
  }

  // ─────────────────────────────────────────────────────────────
  // DELETE DOCUMENT
  // ─────────────────────────────────────────────────────────────

  public async deleteDocument(documentId: string): Promise<boolean> {
    if (this.useMock) {
      await new Promise((resolve) => setTimeout(resolve, 500));
      const index = this.mockDocs.findIndex((d) => d.id === documentId);
      if (index === -1) return false;
      this.mockDocs.splice(index, 1);
      return true;
    }

    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const url = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items(${documentId})`;

      const response = await this.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE",
          },
        },
      );

      return response.ok;
    } catch (error) {
      console.error("DocumentService.deleteDocument:", error);
      return false;
    }
  }

  // ─────────────────────────────────────────────────────────────
  // ARCHIVE DOCUMENT
  // ─────────────────────────────────────────────────────────────

  /**
   * Set document status to 'Archive'.
   * Called as the first step of the Archive & Replace flow.
   * The archive is immediate and irreversible in Release 1.
   */
  public async archiveDocument(documentId: string): Promise<boolean> {
    if (this.useMock) {
      await new Promise((resolve) => setTimeout(resolve, 300));
      const doc = this.mockDocs.find((d) => d.id === documentId);
      if (!doc) return false;
      doc.status = "Archive";
      return true;
    }

    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const url = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items(${documentId})`;

      const response = await this.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Content-Type": "application/json;odata=nometadata",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify({ SpectraStatus: "Archive" }),
        },
      );

      return response.ok;
    } catch (error) {
      console.error("DocumentService.archiveDocument:", error);
      return false;
    }
  }

  public async reActivateDocument(documentId: string): Promise<boolean> {
    if (this.useMock) {
      await new Promise((resolve) => setTimeout(resolve, 300));
      const doc = this.mockDocs.find((d) => d.id === documentId);
      if (!doc) return false;
      const identityKey = this._getDocumentIdentityKeyFromDoc(doc);
      // Archive all Current docs with same identity, then activate the target
      this.mockDocs.forEach((item) => {
        if (item.id === documentId) {
          item.status = "Current";
        } else if (
          item.status === "Current" &&
          this._getDocumentIdentityKeyFromDoc(item) === identityKey
        ) {
          item.status = "Archive";
        }
      });
      return true;
    }

    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${this._documentLibrary}')/items`;

      // Step 1: Fetch the target document's full metadata to build its identity key
      const identityFields = [
        "Id", "FileLeafRef", "FileRef",
        "SpectraDocumentType", "SpectraTherapeuticArea", "SpectraSubTherapeuticArea",
        "SpectraIndication", "SpectraLineOfTherapy", "SpectraAsset",
        "SpectraPAID", "SpectraDiseaseArea", "SpectraStatus",
      ].join(",");

      const targetDocUrl = `${listUrl}(${documentId})?$select=${identityFields}`;
      const targetResponse = await this.context.spHttpClient.get(
        targetDocUrl,
        SPHttpClient.configurations.v1,
      );
      if (!targetResponse.ok) return false;

      const targetItem = await targetResponse.json() as Record<string, unknown>;
      const targetDoc = this._mapSharePointItemToDocument(targetItem);
      const identityKey = this._getDocumentIdentityKeyFromDoc(targetDoc);

      // Step 2: Fetch all Current documents to find any that match the same identity
      const currentDocsUrl = `${listUrl}?$select=${identityFields}&$filter=SpectraStatus eq 'Current'&$top=5000`;
      const currentResponse = await this.context.spHttpClient.get(
        currentDocsUrl,
        SPHttpClient.configurations.v1,
      );
      if (!currentResponse.ok) return false;

      const currentData = await currentResponse.json() as { value: Array<Record<string, unknown>> };
      const docsToArchive = currentData.value
        .map((item) => this._mapSharePointItemToDocument(item))
        .filter((doc) =>
          doc.id !== documentId &&
          this._getDocumentIdentityKeyFromDoc(doc) === identityKey,
        );

      // Step 3: Activate target + archive all matching Current docs in parallel
      const allUpdates = [
        this.context.spHttpClient.post(
          `${listUrl}(${documentId})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              "Content-Type": "application/json;odata=nometadata",
              "IF-MATCH": "*",
              "X-HTTP-Method": "MERGE",
            },
            body: JSON.stringify({ SpectraStatus: "Current" }),
          },
        ),
        ...docsToArchive.map((doc) =>
          this.context.spHttpClient.post(
            `${listUrl}(${doc.id})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                "Content-Type": "application/json;odata=nometadata",
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE",
              },
              body: JSON.stringify({ SpectraStatus: "Archive" }),
            },
          )
        ),
      ];

      const results = await Promise.all(allUpdates);
      return results.every((r) => r.ok);
    } catch (error) {
      console.error("DocumentService.reActivateDocument:", error);
      return false;
    }
  }

  // ─────────────────────────────────────────────────────────────
  // PRIVATE HELPERS
  // ─────────────────────────────────────────────────────────────

  /**
   * Parse a semicolon-separated SharePoint text field into a string array.
   */
  private _parseMultiValue(value: unknown): string[] {
    if (!value || typeof value !== "string") return [];
    return value
      .split(";")
      .map((v: string) => v.trim())
      .filter(Boolean);
  }

  /**
   * Parse a comma-separated search tokens field into a string array.
   */
  private _parseSearchTokens(value: unknown): string[] {
    if (!value || typeof value !== "string") return [];
    return value
      .split(",")
      .map((v: string) => v.trim())
      .filter(Boolean);
  }

  /**
   * Map a SharePoint list item to an IDocument.
   */
  private _mapSharePointItemToDocument(
    item: Record<string, unknown>,
  ): IDocument {
    const fileName = (item.FileLeafRef as string) || "";
    const ext = fileName.split(".").pop() || "";
    const createdBy = (item.Author as Record<string, string>)?.Title || "";
    const createdByEmail = (item.Author as Record<string, string>)?.EMail || "";
    const modifiedBy = (item.Editor as Record<string, string>)?.Title || "";
    const modifiedByEmail =
      (item.Editor as Record<string, string>)?.EMail || "";
    const spectra511 = createdByEmail || createdBy;

    return {
      id: String(item.Id),
      fileName,
      fileExtension: ext,
      fileUrl: (item.FileRef as string) || "",
      fileSize:
        typeof item.File_x0020_Size === "number"
          ? (item.File_x0020_Size as number)
          : 0,
      createdBy,
      createdByEmail,
      modifiedBy,
      modifiedByEmail,
      uploadDate: (item.Created as string) || "",
      title: (item.Title as string) || "",
      asset: this._parseMultiValue(item.SpectraAsset),
      documentType: (item.SpectraDocumentType as string) || "",
      therapeuticArea: this._parseMultiValue(item.SpectraTherapeuticArea),
      subTherapeuticArea: this._parseMultiValue(item.SpectraSubTherapeuticArea),
      indication: this._parseMultiValue(item.SpectraIndication),
      lineOfTherapy: this._parseMultiValue(item.SpectraLineOfTherapy),
      paid: this._parseMultiValue(item.SpectraPAID),
      diseaseArea: this._parseMultiValue(item.SpectraDiseaseArea),
      effectiveDate: (item.SpectraEffectiveDate as string) || "",
      status: ((item.SpectraStatus as string) || "Current") as
        | "Current"
        | "Archive",
      description: (item.SpectraDescription as string) || "",
      comments: this._toTrimmedString(
        item.SpectraComments ?? item.Comments ?? item.comments,
      ),
      spectra511,
      searchTokens: this._parseSearchTokens(item.SpectraSearchTokens),
      immutableFileName: (item.SpectraImmutableFileName as string) || "",
    };
  }

  private _toTrimmedString(value: unknown): string {
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  /**
   * Build a SharePoint metadata update payload from an IUploadPayload.
   */
  private _buildMetadataPayload(
    payload: Partial<IUploadPayload>,
  ): Record<string, unknown> {
    const metadata: Record<string, unknown> = {};

    if (payload.documentType !== undefined)
      metadata.SpectraDocumentType = payload.documentType;
    if (payload.therapeuticArea !== undefined)
      metadata.SpectraTherapeuticArea = payload.therapeuticArea.join(
        MULTI_VALUE_SEPARATOR,
      );
    if (payload.subTherapeuticArea !== undefined)
      metadata.SpectraSubTherapeuticArea = payload.subTherapeuticArea.join(
        MULTI_VALUE_SEPARATOR,
      );
    if (payload.asset !== undefined)
      metadata.SpectraAsset = payload.asset.join(MULTI_VALUE_SEPARATOR);
    if (payload.indication !== undefined)
      metadata.SpectraIndication = payload.indication.join(
        MULTI_VALUE_SEPARATOR,
      );
    if (payload.lineOfTherapy !== undefined)
      metadata.SpectraLineOfTherapy = payload.lineOfTherapy.join(
        MULTI_VALUE_SEPARATOR,
      );
    if (payload.paid !== undefined)
      metadata.SpectraPAID = payload.paid.join(MULTI_VALUE_SEPARATOR);
    if (payload.diseaseArea !== undefined)
      metadata.SpectraDiseaseArea = payload.diseaseArea.join(
        MULTI_VALUE_SEPARATOR,
      );
    if (payload.effectiveDate !== undefined)
      metadata.SpectraEffectiveDate = payload.effectiveDate;
    if (payload.description !== undefined)
      metadata.SpectraDescription = payload.description;
    if (payload.comments !== undefined)
      metadata.SpectraComments = payload.comments;
    if (payload.searchTokens !== undefined)
      metadata.SpectraSearchTokens = payload.searchTokens.join(
        SEARCH_TOKENS_SEPARATOR,
      );

    metadata.SpectraStatus = "Current";
    // SpectraImmutableFileName is set by the upload flow with the original filename

    return metadata;
  }
}
