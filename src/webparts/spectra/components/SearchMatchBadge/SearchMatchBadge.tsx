import * as React from "react";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { SearchMatchKind } from "../../utils/filterHelper";

export interface ISearchMatchBadgeProps {
  kind: SearchMatchKind;
}

export const SearchMatchBadge: React.FC<ISearchMatchBadgeProps> = ({ kind }) => {
  if (kind !== "exact") return null;

  return (
    <TooltipHost content="Exact Match" styles={{ root: { display: "flex", alignItems: "center", flexShrink: 0 } }}>
      <img
        src={require("../../assets/icons/check-circle.svg")}
        alt="Exact match"
        style={{ width: 17, height: 16, display: "block" }}
        aria-label="Exact match"
      />
    </TooltipHost>
  );
};
