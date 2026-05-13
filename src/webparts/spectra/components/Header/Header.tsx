import * as React from "react";
import { EffectiveRole } from "../../interfaces/IAuthResponse";
import { RoleBadge } from "../RoleBadge/RoleBadge";
import { Icon } from "@fluentui/react/lib/Icon";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import styles from "../SPECTRA.module.scss";

export interface IHeaderProps {
  role: EffectiveRole;
  userDisplayName: string;
  onLogoClick: () => void;
  enableDevRoleSwitch: boolean;
  onRoleBadgeClick: () => void;
}

export const Header: React.FC<IHeaderProps> = ({
  role,
  userDisplayName,
  onLogoClick,
  enableDevRoleSwitch,
  onRoleBadgeClick,
}) => {
  return (
    <header className={styles.header}>
      <div
        className={styles.headerLogo}
        onClick={onLogoClick}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => {
          if (e.key === "Enter") onLogoClick();
        }}
        aria-label="PEAKS home"
      >
        <div className={styles.headerLogoIcon}>a</div>
        <span className={styles.headerLogoText}>PEAKS</span>
      </div>

      <div className={styles.headerRight}>
        <TooltipHost content="Help">
          <Icon iconName="Help" className={styles.headerHelpIcon} />
        </TooltipHost>
        <TooltipHost content="Ask R&D">
          <a className={styles.headerAskRd} href="#" aria-label="Ask R&D">
            <span>Ask R&D</span>
            <Icon iconName="OpenInNewWindow" />
          </a>
        </TooltipHost>
        <RoleBadge
          role={role}
          isInteractive={enableDevRoleSwitch}
          onClick={onRoleBadgeClick}
        />
        <TooltipHost content={userDisplayName}>
          <div
            className={styles.headerAvatar}
            aria-label={`User: ${userDisplayName}`}
          >
            <Icon iconName="Contact" className={styles.headerAvatarIcon} />
          </div>
        </TooltipHost>
      </div>
    </header>
  );
};
