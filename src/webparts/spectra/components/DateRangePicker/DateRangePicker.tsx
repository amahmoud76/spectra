import * as React from "react";
import { DatePicker, IDatePickerStrings } from "@fluentui/react/lib/DatePicker";
import { DayOfWeek } from "@fluentui/react/lib/Calendar";
import styles from "../SPECTRA.module.scss";

const pad2 = (n: number): string => (n < 10 ? "0" + n : "" + n);

const dateStrings: IDatePickerStrings = {
  months: [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ],
  shortMonths: [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ],
  days: [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ],
  shortDays: ["S", "M", "T", "W", "T", "F", "S"],
  goToToday: "Go to today",
  prevMonthAriaLabel: "Previous month",
  nextMonthAriaLabel: "Next month",
  prevYearAriaLabel: "Previous year",
  nextYearAriaLabel: "Next year",
  closeButtonAriaLabel: "Close",
  isRequiredErrorMessage: "Date is required",
  invalidInputErrorMessage: "Invalid date format",
};

const calendarDayStyles = {
  daySelected: {
    backgroundColor: "#0066F5",
    color: "#ffffff",
    selectors: {
      "&:hover": { backgroundColor: "#0052CC", color: "#ffffff" },
      "& button": { color: "#ffffff" },
    },
  },
  dayIsToday: {
    backgroundColor: "#0066F5",
    color: "#ffffff",
    selectors: { "& button": { color: "#ffffff" } },
  },
};

const inlineDatePickerStyles = {
  root: { flex: 1, minWidth: 0 },
  textField: {
    selectors: {
      ".ms-TextField-fieldGroup": {
        border: "none",
        background: "transparent",
        minHeight: "unset",
        height: "28px",
      },
      ".ms-TextField-field": {
        fontSize: "13px",
        padding: "0 4px",
        color: "#1f2937",
        background: "transparent",
      },
      ".ms-TextField-fieldGroup:hover": { border: "none" },
      ".ms-TextField-fieldGroup:focus-within": { border: "none" },
    },
  },
  callout: { minWidth: 280 },
};

export interface IDateRangePickerProps {
  label: string;
  fromDate: Date | null;
  toDate: Date | null;
  onFromChange: (date: Date | null) => void;
  onToChange: (date: Date | null) => void;
}

export const DateRangePicker: React.FC<IDateRangePickerProps> = ({
  label,
  fromDate,
  toDate,
  onFromChange,
  onToChange,
}) => {
  return (
    <div className={styles.dateRangePickerWrapper}>
      <span className={styles.dateRangePickerLabel}>{label}</span>
      <div className={styles.dateRangePickerInner}>
        <DatePicker
          value={fromDate || undefined}
          onSelectDate={(d) => onFromChange(d || null)}
          placeholder="DD/MMM/YYYY"
          firstDayOfWeek={DayOfWeek.Monday}
          strings={dateStrings}
          allowTextInput={true}
          formatDate={(d) => {
            if (!d) return "";
            const day = pad2(d.getDate());
            const month = dateStrings.shortMonths![d.getMonth()];
            const year = d.getFullYear();
            return `${day}/${month}/${year}`;
          }}
          parseDateFromString={(s) => {
            const parts = s.split("/");
            if (parts.length === 3) {
              const day = parseInt(parts[0], 10);
              const monthIdx = dateStrings.shortMonths!.findIndex(
                (m) => m.toLowerCase() === parts[1].toLowerCase(),
              );
              const year = parseInt(parts[2], 10);
              if (!isNaN(day) && monthIdx >= 0 && !isNaN(year)) {
                return new Date(year, monthIdx, day);
              }
            }
            return null;
          }}
          calendarProps={{ calendarDayProps: { styles: calendarDayStyles } }}
          styles={inlineDatePickerStyles}
          ariaLabel={`${label} start date`}
        />
        <span className={styles.dateRangePickerSeparator}>–</span>
        <DatePicker
          value={toDate || undefined}
          onSelectDate={(d) => onToChange(d || null)}
          placeholder="DD/MMM/YYYY"
          firstDayOfWeek={DayOfWeek.Monday}
          strings={dateStrings}
          allowTextInput={true}
          minDate={fromDate || undefined}
          formatDate={(d) => {
            if (!d) return "";
            const day = pad2(d.getDate());
            const month = dateStrings.shortMonths![d.getMonth()];
            const year = d.getFullYear();
            return `${day}/${month}/${year}`;
          }}
          parseDateFromString={(s) => {
            const parts = s.split("/");
            if (parts.length === 3) {
              const day = parseInt(parts[0], 10);
              const monthIdx = dateStrings.shortMonths!.findIndex(
                (m) => m.toLowerCase() === parts[1].toLowerCase(),
              );
              const year = parseInt(parts[2], 10);
              if (!isNaN(day) && monthIdx >= 0 && !isNaN(year)) {
                return new Date(year, monthIdx, day);
              }
            }
            return null;
          }}
          calendarProps={{ calendarDayProps: { styles: calendarDayStyles } }}
          styles={inlineDatePickerStyles}
          ariaLabel={`${label} end date`}
        />
      </div>
    </div>
  );
};
