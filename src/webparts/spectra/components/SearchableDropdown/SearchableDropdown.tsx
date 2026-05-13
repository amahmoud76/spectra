import * as React from "react";
import {
  TagPicker,
  ITag,
  IBasePickerSuggestionsProps,
} from "@fluentui/react/lib/Pickers";
import { Label } from "@fluentui/react/lib/Label";
import styles from "../SPECTRA.module.scss";

export interface ISearchableDropdownProps {
  label: string;
  required?: boolean;
  options: string[];
  selectedKeys: string[];
  onChange: (selected: string[]) => void;
  placeholder?: string;
  multiSelect?: boolean;
  disabled?: boolean;
  showChipsBelow?: boolean;
  errorMessage?: string;
}

const normalizeText = (value: unknown): string =>
  typeof value === "string" ? value.trim() : "";

const ensureArray = (value: string[] | string | undefined | null): string[] => {
  const values = Array.isArray(value)
    ? value
    : typeof value === "string"
      ? [value]
      : [];

  return values
    .map((item) => normalizeText(item))
    .filter((item) => item.length > 0)
    .filter((item, index, items) => items.indexOf(item) === index);
};

// ── Single-Select — uses TagPicker with key reset ─────────────
const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "Suggestions",
  noResultsFoundText: "No matches found",
};

const SingleSelectPicker: React.FC<{
  allTags: ITag[];
  selectedKeys: string[];
  onChange: (selected: string[]) => void;
  placeholder: string;
  label: string;
  disabled: boolean;
  showChipsBelow?: boolean;
}> = ({
  allTags,
  selectedKeys: rawSelectedKeys,
  onChange,
  placeholder,
  label,
  disabled,
  showChipsBelow = false,
}) => {
  const selectedKeys = React.useMemo(
    () => ensureArray(rawSelectedKeys),
    [rawSelectedKeys],
  );

  // If showChipsBelow, use MultiSelectNative rendering but enforce single selection
  if (showChipsBelow) {
    // eslint-disable-next-line @typescript-eslint/no-use-before-define
    const allOptions = allTags.map((t) => t.name);
    // eslint-disable-next-line @typescript-eslint/no-use-before-define
    return (
      <MultiSelectNativeChipsSingle
        options={allOptions}
        selectedKeys={selectedKeys}
        onChange={onChange}
        placeholder={placeholder}
        label={label}
        disabled={disabled}
      />
    );
  }

  // Default: use TagPicker with inline rendering
  const [pickerKey, setPickerKey] = React.useState(0);

  const selectedKeySet = React.useMemo(
    () => new Set(selectedKeys),
    [selectedKeys],
  );

  const selectedTags: ITag[] = React.useMemo(
    () => selectedKeys.map((key) => ({ key, name: key })),
    [selectedKeys],
  );

  const availableTags = React.useMemo(
    () => allTags.filter((tag) => !selectedKeySet.has(tag.key as string)),
    [allTags, selectedKeySet],
  );

  const onResolveSuggestions = React.useCallback(
    (filterText: string): ITag[] => {
      if (!filterText) return availableTags;
      const lower = filterText.toLowerCase();
      return availableTags.filter(
        (t) => t.name && t.name.toLowerCase().includes(lower),
      );
    },
    [availableTags],
  );

  const onEmptyResolveSuggestions = React.useCallback(
    (): ITag[] => availableTags,
    [availableTags],
  );

  const onTagsChanged = React.useCallback(
    (items?: ITag[]) => {
      if (!items || items.length === 0) {
        onChange([]);
      } else {
        const lastItem = items[items.length - 1];
        onChange([lastItem.key as string]);
      }
      setPickerKey((prev) => prev + 1);
    },
    [onChange],
  );

  return (
    <div className={styles.searchPickerWrap}>
      <img src={require('../../assets/icons/search.svg')} alt="" className={styles.searchPickerIcon} aria-hidden="true" style={{ width: '16px', height: '16px', display: 'block' }} />
      <TagPicker
        key={pickerKey}
        onResolveSuggestions={onResolveSuggestions}
        onEmptyResolveSuggestions={onEmptyResolveSuggestions}
        defaultSelectedItems={selectedTags}
        onChange={onTagsChanged}
        pickerSuggestionsProps={suggestionProps}
        inputProps={{
          placeholder: selectedKeys.length === 0 ? placeholder : "",
          "aria-label": label,
        }}
        disabled={disabled}
        resolveDelay={100}
      />
    </div>
  );
};

// ── Single-Select with Chips Below ────────────────────────────
// Uses custom HTML + filtered dropdown + chips below (but enforces single selection)
const MultiSelectNativeChipsSingle: React.FC<{
  options: string[];
  selectedKeys: string[];
  onChange: (selected: string[]) => void;
  placeholder: string;
  label: string;
  disabled: boolean;
}> = ({
  options,
  selectedKeys: rawSelectedKeys,
  onChange,
  placeholder,
  label,
  disabled,
}) => {
  const [searchText, setSearchText] = React.useState("");
  const [isOpen, setIsOpen] = React.useState(false);
  const [highlightIndex, setHighlightIndex] = React.useState(-1);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const inputRef = React.useRef<HTMLInputElement>(null);
  const listRef = React.useRef<HTMLDivElement>(null);

  const selectedKeys = ensureArray(rawSelectedKeys);
  const selectedKeySet = new Set(selectedKeys);

  // Filter options: remove already selected, then filter by search text
  const filteredOptions = React.useMemo(() => {
    const available = options.filter((opt) => !selectedKeySet.has(opt));
    if (!searchText) return available;
    const lower = searchText.toLowerCase();
    return available.filter((opt) => opt && opt.toLowerCase().includes(lower));
  }, [options, selectedKeySet, searchText]);

  // Select an item (single selection)
  const selectItem = (value: string): void => {
    const sanitizedValue = normalizeText(value);
    if (!sanitizedValue) return;

    onChange([sanitizedValue]); // Replace, don't append
    setSearchText("");
    setIsOpen(true);
    setHighlightIndex(-1);
    if (inputRef.current) {
      inputRef.current.focus();
    }
  };

  // Remove the item (since single-select, removes the only selection)
  const removeItem = (): void => {
    onChange([]);
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    setSearchText(e.target.value);
    setIsOpen(true);
    setHighlightIndex(-1);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === "Escape") {
      setIsOpen(false);
      setHighlightIndex(-1);
      return;
    }

    if (
      e.key === "Backspace" &&
      searchText === "" &&
      selectedKeys.length > 0 &&
      !disabled
    ) {
      onChange([]);
      return;
    }

    if (e.key === "ArrowDown") {
      e.preventDefault();
      setIsOpen(true);
      setHighlightIndex((prev) => {
        const next = prev < filteredOptions.length - 1 ? prev + 1 : 0;
        scrollToIndex(next);
        return next;
      });
      return;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      setHighlightIndex((prev) => {
        const next = prev > 0 ? prev - 1 : filteredOptions.length - 1;
        scrollToIndex(next);
        return next;
      });
      return;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      if (highlightIndex >= 0 && highlightIndex < filteredOptions.length) {
        selectItem(filteredOptions[highlightIndex]);
      }
      return;
    }
  };

  const scrollToIndex = (index: number): void => {
    if (listRef.current && listRef.current.children[index]) {
      (listRef.current.children[index] as HTMLElement).scrollIntoView({
        block: "nearest",
      });
    }
  };

  const handleFocus = (): void => {
    setIsOpen(true);
  };

  const handleClick = (): void => {
    setIsOpen(true);
  };

  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (
        containerRef.current &&
        !containerRef.current.contains(e.target as Node)
      ) {
        setIsOpen(false);
        setHighlightIndex(-1);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  return (
    <div ref={containerRef} className={styles.nativeMultiSelect}>
      <div className={styles.nativeMultiSelectInputWrap}>
        <img src={require('../../assets/icons/search.svg')} alt="" className={styles.searchPickerIcon} aria-hidden="true" style={{ width: '16px', height: '16px', display: 'block' }} />
        <input
          ref={inputRef}
          type="text"
          className={styles.nativeMultiSelectInput}
          value={searchText}
          onChange={handleInputChange}
          onFocus={handleFocus}
          onClick={handleClick}
          onKeyDown={handleKeyDown}
          placeholder={placeholder}
          aria-label={label}
          aria-expanded={isOpen}
          aria-haspopup="listbox"
          role="combobox"
          autoComplete="off"
          disabled={disabled}
        />
      </div>

      {isOpen && filteredOptions.length > 0 && (
        <div
          ref={listRef}
          className={styles.nativeMultiSelectDropdown}
          role="listbox"
        >
          {filteredOptions.map((opt, idx) => (
            <div
              key={opt}
              className={
                styles.nativeMultiSelectOption +
                (idx === highlightIndex
                  ? " " + styles.nativeMultiSelectOptionHighlight
                  : "")
              }
              onMouseDown={(e) => {
                e.preventDefault();
                selectItem(opt);
              }}
              onMouseEnter={() => setHighlightIndex(idx)}
              role="option"
              aria-selected={idx === highlightIndex}
            >
              {opt}
            </div>
          ))}
        </div>
      )}

      {isOpen && filteredOptions.length === 0 && searchText && (
        <div className={styles.nativeMultiSelectDropdown}>
          <div className={styles.nativeMultiSelectNoResults}>
            No matches found
          </div>
        </div>
      )}

      {selectedKeys.length > 0 && (
        <div className={styles.nativeMultiSelectChips}>
          {selectedKeys.map((key) => (
            <span key={key} className={styles.chip}>
              {key}
              <button
                className={styles.chipRemove}
                onClick={() => !disabled && removeItem()}
                aria-label={`Remove ${key}`}
                type="button"
                disabled={disabled}
              >
                ×
              </button>
            </span>
          ))}
        </div>
      )}
    </div>
  );
};

// ── Multi-Select — Pure HTML/React (no TagPicker) ─────────────
// Uses a plain input + filtered dropdown + chips below.
// No Fluent UI TagPicker involved — avoids all its bugs.
const MultiSelectNative: React.FC<{
  options: string[];
  selectedKeys: string[];
  onChange: (selected: string[]) => void;
  placeholder: string;
  label: string;
  disabled: boolean;
}> = ({
  options,
  selectedKeys: rawSelectedKeys,
  onChange,
  placeholder,
  label,
  disabled,
}) => {
  const [searchText, setSearchText] = React.useState("");
  const [isOpen, setIsOpen] = React.useState(false);
  const [highlightIndex, setHighlightIndex] = React.useState(-1);
  const containerRef = React.useRef<HTMLDivElement>(null);
  const inputRef = React.useRef<HTMLInputElement>(null);
  const listRef = React.useRef<HTMLDivElement>(null);

  const selectedKeys = ensureArray(rawSelectedKeys);
  const selectedKeySet = new Set(selectedKeys);

  // Filter options: remove already selected, then filter by search text
  const filteredOptions = React.useMemo(() => {
    const available = options.filter((opt) => !selectedKeySet.has(opt));
    if (!searchText) return available;
    const lower = searchText.toLowerCase();
    return available.filter((opt) => opt && opt.toLowerCase().includes(lower));
  }, [options, selectedKeySet, searchText]);

  // Select an item
  const selectItem = (value: string): void => {
    const sanitizedValue = normalizeText(value);
    if (!sanitizedValue) return;

    const current = ensureArray(rawSelectedKeys);
    if (!current.includes(sanitizedValue)) {
      onChange([...current, sanitizedValue]);
    }
    setSearchText("");
    setIsOpen(true);
    setHighlightIndex(-1);
    // Keep focus on input for next selection
    if (inputRef.current) {
      inputRef.current.focus();
    }
  };

  // Remove a chip
  const removeItem = (value: string): void => {
    const current = ensureArray(rawSelectedKeys);
    onChange(current.filter((k) => k !== value));
  };

  // Handle typing in the search input
  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    setSearchText(e.target.value);
    setIsOpen(true);
    setHighlightIndex(-1);
  };

  // Handle keyboard
  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === "Escape") {
      setIsOpen(false);
      setHighlightIndex(-1);
      return;
    }

    if (e.key === "Backspace" && searchText === "" && selectedKeys.length > 0) {
      const current = ensureArray(rawSelectedKeys);
      onChange(current.slice(0, -1));
      return;
    }

    if (e.key === "ArrowDown") {
      e.preventDefault();
      setIsOpen(true);
      setHighlightIndex((prev) => {
        const next = prev < filteredOptions.length - 1 ? prev + 1 : 0;
        scrollToIndex(next);
        return next;
      });
      return;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      setHighlightIndex((prev) => {
        const next = prev > 0 ? prev - 1 : filteredOptions.length - 1;
        scrollToIndex(next);
        return next;
      });
      return;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      if (highlightIndex >= 0 && highlightIndex < filteredOptions.length) {
        selectItem(filteredOptions[highlightIndex]);
      }
      return;
    }
  };

  const scrollToIndex = (index: number): void => {
    if (listRef.current && listRef.current.children[index]) {
      (listRef.current.children[index] as HTMLElement).scrollIntoView({
        block: "nearest",
      });
    }
  };

  // Open on focus
  const handleFocus = (): void => {
    setIsOpen(true);
  };

  const handleClick = (): void => {
    setIsOpen(true);
  };

  // Close when clicking outside
  React.useEffect(() => {
    const handler = (e: MouseEvent): void => {
      if (
        containerRef.current &&
        !containerRef.current.contains(e.target as Node)
      ) {
        setIsOpen(false);
        setHighlightIndex(-1);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  return (
    <div ref={containerRef} className={styles.nativeMultiSelect}>
      {/* Search input */}
      <div className={styles.nativeMultiSelectInputWrap}>
        <img src={require('../../assets/icons/search.svg')} alt="" className={styles.searchPickerIcon} aria-hidden="true" style={{ width: '16px', height: '16px', display: 'block' }} />
        <input
          ref={inputRef}
          type="text"
          className={styles.nativeMultiSelectInput}
          value={searchText}
          onChange={handleInputChange}
          onFocus={handleFocus}
          onClick={handleClick}
          onKeyDown={handleKeyDown}
          placeholder={placeholder}
          aria-label={label}
          aria-expanded={isOpen}
          aria-haspopup="listbox"
          role="combobox"
          autoComplete="off"
          disabled={disabled}
        />
      </div>

      {/* Dropdown suggestions */}
      {isOpen && filteredOptions.length > 0 && (
        <div
          ref={listRef}
          className={styles.nativeMultiSelectDropdown}
          role="listbox"
        >
          {filteredOptions.map((opt, idx) => (
            <div
              key={opt}
              className={
                styles.nativeMultiSelectOption +
                (idx === highlightIndex
                  ? " " + styles.nativeMultiSelectOptionHighlight
                  : "")
              }
              onMouseDown={(e) => {
                e.preventDefault();
                selectItem(opt);
              }}
              onMouseEnter={() => setHighlightIndex(idx)}
              role="option"
              aria-selected={idx === highlightIndex}
            >
              {opt}
            </div>
          ))}
        </div>
      )}

      {isOpen && filteredOptions.length === 0 && searchText && (
        <div className={styles.nativeMultiSelectDropdown}>
          <div className={styles.nativeMultiSelectNoResults}>
            No matches found
          </div>
        </div>
      )}

      {/* Selected chips below the input */}
      {selectedKeys.length > 0 && (
        <div className={styles.nativeMultiSelectChips}>
          {selectedKeys.map((key) => (
            <span key={key} className={styles.chip}>
              {key}
              <button
                className={styles.chipRemove}
                onClick={() => removeItem(key)}
                aria-label={`Remove ${key}`}
                type="button"
              >
                ×
              </button>
            </span>
          ))}
        </div>
      )}
    </div>
  );
};

// ── Main Component ────────────────────────────────────────────
export const SearchableDropdown: React.FC<ISearchableDropdownProps> = ({
  label,
  required = false,
  options,
  selectedKeys,
  onChange,
  placeholder = "Type to search...",
  multiSelect = true,
  disabled = false,
  showChipsBelow = false,
  errorMessage,
}) => {
  const sanitizedOptions = React.useMemo(
    () =>
      (options as unknown[])
        .map((option) => normalizeText(option))
        .filter((option) => option.length > 0)
        .filter((option, index, values) => values.indexOf(option) === index),
    [options],
  );

  const allTags: ITag[] = React.useMemo(
    () => sanitizedOptions.map((opt) => ({ key: opt, name: opt })),
    [sanitizedOptions],
  );

  const safeSelectedKeys = React.useMemo(
    () => ensureArray(selectedKeys),
    [selectedKeys],
  );

  return (
    <div className={styles.formGroup}>
      <Label required={required}>{label}</Label>
      {multiSelect ? (
        <MultiSelectNative
          options={sanitizedOptions}
          selectedKeys={safeSelectedKeys}
          onChange={onChange}
          placeholder={placeholder}
          label={label}
          disabled={disabled}
        />
      ) : (
        <SingleSelectPicker
          allTags={allTags}
          selectedKeys={safeSelectedKeys}
          onChange={onChange}
          placeholder={placeholder}
          label={label}
          disabled={disabled}
          showChipsBelow={showChipsBelow}
        />
      )}
      {errorMessage && <div className={styles.formErrorText}>{errorMessage}</div>}
    </div>
  );
};
