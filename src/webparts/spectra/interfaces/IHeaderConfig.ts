export interface IHeaderLink {
  id: string;
  label: string;
  url: string;
  openInNewTab: boolean;
  sortOrder: number;
  isActive: boolean;
  iconName?: string;
}

export interface IHeaderConfig {
  parentAppName: string;
  parentAppUrl: string;
  parentLogoUrl?: string;
  activeAppLabel: string;
  activeAppUrl?: string;
  activeAppIconName?: string;
  links: IHeaderLink[];
}

export const FALLBACK_HEADER_CONFIG: IHeaderConfig = {
  parentAppName: "Illumination Hub",
  parentAppUrl: "#",
  parentLogoUrl: "https://abbvie.sharepoint.com/teams/SPECTRA/SiteAssets/Logo.png",
  activeAppLabel: "SPECTRA",
  activeAppUrl: "#",
  activeAppIconName: "fa-regular fa-rainbow-half",
  links: [
    {
      id: "astra",
      label: "ASTRA",
      url: "#",
      openInNewTab: true,
      sortOrder: 10,
      isActive: true,
      iconName: "fa-arrow-up-right-from-square",
    },
    {
      id: "inform",
      label: "INFORM",
      url: "#",
      openInNewTab: true,
      sortOrder: 20,
      isActive: true,
      iconName: "fa-arrow-up-right-from-square",
    },
  ],
};
