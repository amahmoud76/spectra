import * as React from "react";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { EffectiveRole } from "../../interfaces/IAuthResponse";
import { IHeaderConfig } from "../../interfaces/IHeaderConfig";
import { RoleBadge } from "../RoleBadge/RoleBadge";
import styles from "../SPECTRA.module.scss";

export interface IParentHeaderProps {
  config: IHeaderConfig;
  role: EffectiveRole;
  userDisplayName: string;
  userEmail: string;
  siteUrl: string;
  enableDevRoleSwitch: boolean;
  onRoleBadgeClick: () => void;
  onSpectraClick: () => void;
}

export const ParentHeader: React.FC<IParentHeaderProps> = ({
  config,
  role,
  userDisplayName,
  userEmail,
  siteUrl,
  enableDevRoleSwitch,
  onRoleBadgeClick,
  onSpectraClick,
}) => {
  const styleClassMap = styles as unknown as Record<string, string>;
  const activeUrl = config.activeAppUrl || "#";
  const [imgError, setImgError] = React.useState(false);
  const [userPhotoError, setUserPhotoError] = React.useState(false);
  const userPhotoUrl = React.useMemo(() => {
    if (!userEmail) return "";
    return `${siteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${encodeURIComponent(userEmail)}`;
  }, [siteUrl, userEmail]);
  const userInitials =
    userDisplayName
      .trim()
      .split(/\s+/)
      .slice(0, 2)
      .map((part) => part[0]?.toUpperCase() || "")
      .join("") || "U";

  const getLinkIdentity = React.useCallback(
    (link: { id: string; label: string }) => {
      const normalized = `${link.id} ${link.label}`.toLowerCase();
      if (normalized.includes("astra")) return "astra";
      if (normalized.includes("inform")) return "inform";
      return "other";
    },
    [],
  );

  const astraLink = React.useMemo(
    () => config.links.find((link) => getLinkIdentity(link) === "astra"),
    [config.links, getLinkIdentity],
  );

  const informLink = React.useMemo(
    () => config.links.find((link) => getLinkIdentity(link) === "inform"),
    [config.links, getLinkIdentity],
  );

  const otherLinks = React.useMemo(
    () =>
      config.links.filter(
        (link) =>
          getLinkIdentity(link) !== "astra" &&
          getLinkIdentity(link) !== "inform",
      ),
    [config.links, getLinkIdentity],
  );

  const renderNavLink = React.useCallback(
    (link: {
      id: string;
      label: string;
      url: string;
      openInNewTab?: boolean;
      iconName?: string;
    }) => (
      <a
        key={link.id}
        href={link.url}
        className={styles.parentHeaderNavLink}
        target={link.openInNewTab ? "_blank" : "_self"}
        rel="noreferrer"
      >
        {getLinkIdentity(link) === "astra" && (
          <img src={require('../../assets/icons/astra.svg')} alt="" style={{ width: '16px', height: '16px', display: 'inline-block', marginRight: '4px' }} className={styles.parentHeaderNavLinkIcon} aria-hidden="true" />
        )}
        {link.label}
        {(link.openInNewTab || getLinkIdentity(link) === "inform") && (
          <img src={require('../../assets/icons/external.svg')} alt="" style={{ width: '16px', height: '16px', display: 'inline-block', marginLeft: '4px' }} className={styles.parentHeaderNavLinkIcon} aria-hidden="true" />
        )}
      </a>
    ),
    [getLinkIdentity],
  );

  React.useEffect(() => {
    setImgError(false);
  }, [config.parentLogoUrl]);

  React.useEffect(() => {
    setUserPhotoError(false);
  }, [userPhotoUrl]);

  return (
    <header className={styles.parentHeader}>
      <div className={styles.parentHeaderBrandWrap}>
        <a
          className={styles.parentHeaderBrand}
          href={config.parentAppUrl}
          target="_self"
          rel="noreferrer"
          aria-label={`${config.parentAppName} home`}
        >
          {config.parentLogoUrl && !imgError ? (
            <img
              src={config.parentLogoUrl}
              alt=""
              className={styles.parentHeaderBrandImage}
              onError={() => setImgError(true)}
            />
          ) : (
            <div className={styles.parentHeaderBrandIcon}>i</div>
          )}
          <span className={styles.parentHeaderBrandText}>
            {config.parentAppName}
          </span>
        </a>
      </div>

      <nav className={styles.parentHeaderNav} aria-label="Product navigation">
        {astraLink && renderNavLink(astraLink)}

        <a
          href={activeUrl}
          className={`${styles.parentHeaderNavLink} ${styles.parentHeaderNavLinkActive}`}
          onClick={(event) => {
            event.preventDefault();
            onSpectraClick();
          }}
          aria-current="page"
        >
          <img
            src={require('../../assets/icons/rainbow-half.svg')}
            alt=""
            style={{ display: 'inline-block', width: '1em', height: '1em', marginRight: '4px' }}
            aria-hidden="true"
          />
          {config.activeAppLabel}
        </a>

        {informLink && renderNavLink(informLink)}
        {otherLinks.map((link) => renderNavLink(link))}
      </nav>

      <div className={styles.parentHeaderRight}>
        <TooltipHost
          content={
            <span className={styleClassMap.parentHeaderSupportTooltip}>
              <img
                src={require("../../assets/icons/email.svg")}
                alt=""
                className={styleClassMap.parentHeaderSupportTooltipIcon}
                aria-hidden="true"
              />
              <span className={styleClassMap.parentHeaderSupportTooltipText}>
                Contact SPECTRA Support
              </span>
            </span>
          }
          calloutProps={{ gapSpace: 8 }}
        >
          <a
            className={styleClassMap.parentHeaderSupportLink}
            href="mailto:SPECTRAsupport@abbvie.com"
            aria-label="Email SPECTRA Support"
            title="Email SPECTRA Support"
          >
            <img
              src={require('../../assets/icons/support.svg')}
              alt=""
              className={styleClassMap.parentHeaderSupportIcon}
              aria-hidden="true"
            />
          </a>
        </TooltipHost>

        <RoleBadge
          role={role}
          isInteractive={enableDevRoleSwitch}
          onClick={onRoleBadgeClick}
        />

        <TooltipHost content={userDisplayName}>
          <div
            className={styles.parentHeaderAvatar}
            aria-label={`User: ${userDisplayName}`}
          >
            {userPhotoUrl && !userPhotoError ? (
              <img
                className={styles.parentHeaderAvatarImage}
                src={userPhotoUrl}
                alt=""
                onError={() => setUserPhotoError(true)}
              />
            ) : (
              <span className={styles.parentHeaderAvatarInitials}>
                {userInitials}
              </span>
            )}
          </div>
        </TooltipHost>
      </div>
    </header>
  );
};
