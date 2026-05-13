import * as React from "react";
import { EffectiveRole } from "../../interfaces/IAuthResponse";
import styles from "../SPECTRA.module.scss";

export interface IRoleBadgeProps {
  role: EffectiveRole;
  isInteractive?: boolean;
  onClick?: () => void;
}

const ROLE_CONFIG: Record<
  EffectiveRole,
  { label: string; className: string; icon: string }
> = {
  admin: {
    label: "Administrator",
    className: styles.roleBadgeAdmin,
    icon: require("../../assets/icons/administrator.svg"),
  },
  contributor: {
    label: "Contributor",
    className: styles.roleBadgeContributor,
    icon: require("../../assets/icons/contributor.svg"),
  },
  viewer: {
    label: "View Only",
    className: styles.roleBadgeViewer,
    icon: require("../../assets/icons/viewonly.svg"),
  },
};

export const RoleBadge: React.FC<IRoleBadgeProps> = ({
  role,
  isInteractive = false,
  onClick,
}) => {
  const config = ROLE_CONFIG[role];

  return (
    <span
      className={`${styles.roleBadge} ${config.className} ${isInteractive ? styles.roleBadgeInteractive : ""}`}
      aria-label={
        isInteractive
          ? `Role: ${config.label}. Click to switch role in dev mode`
          : `Role: ${config.label}`
      }
      role={isInteractive ? "button" : undefined}
      tabIndex={isInteractive ? 0 : undefined}
      onClick={isInteractive ? onClick : undefined}
      onKeyDown={
        isInteractive
          ? (e) => {
              if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                onClick?.();
              }
            }
          : undefined
      }
    >
      <img src={config.icon} alt="" style={{ width: '12px', height: '12px', display: 'inline-block', marginRight: '4px' }} aria-hidden="true" />
      {config.label}
    </span>
  );
};
