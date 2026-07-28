import * as React from "react";
import { IDocument } from "../../interfaces/IDocument";
import styles from "../SPECTRA.module.scss";

export interface ISearchSuggestionsProps {
  documents: IDocument[];
  isLoading?: boolean;
  onSelect: (doc: IDocument) => void;
}

const buildMeta = (doc: IDocument): string =>
  [doc.documentType, ...doc.asset].filter(Boolean).join(" · ");

export const SearchSuggestions: React.FC<ISearchSuggestionsProps> = ({
  documents,
  isLoading = false,
  onSelect,
}) => {
  if (isLoading) {
    return (
      <div className={styles.searchSuggestions}>
        <div className={styles.searchSuggestionEmpty}>Searching…</div>
      </div>
    );
  }

  if (documents.length === 0) {
    return (
      <div className={styles.searchSuggestions}>
        <div className={styles.searchSuggestionEmpty}>No matches found</div>
      </div>
    );
  }

  return (
    <div className={styles.searchSuggestions} role="listbox">
      {documents.map((doc) => (
        <div
          key={doc.id}
          className={styles.searchSuggestionItem}
          role="option"
          aria-selected={false}
          // onMouseDown (not onClick) so the suggestion is selected before
          // the input's onBlur closes this list.
          onMouseDown={(e) => {
            e.preventDefault();
            onSelect(doc);
          }}
        >
          <div className={styles.searchSuggestionTitle}>{doc.fileName}</div>
          <div className={styles.searchSuggestionMeta}>{buildMeta(doc)}</div>
        </div>
      ))}
    </div>
  );
};
