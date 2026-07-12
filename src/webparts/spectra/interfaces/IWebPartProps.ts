import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IWebPartProps {
  context: WebPartContext; // SPFx context — provides user info, HTTP clients, page context
  pageSize: number; // Documents per page — configured in property pane
  title: string; // Web part title from property pane
  userEmail: string; // From context.pageContext.user.email
  documentLibrary: string; // Target document library — configured in property pane
  inactivityTimeoutMinutes: number; // Session lock timeout after inactivity
  useMock: boolean; // Runtime mode toggle: mock/demo vs live SharePoint
  useAdGroups: boolean; // Runtime auth mode toggle: AD groups vs mock role
  mockRole: "admin" | "contributor" | "viewer"; // Fallback role when not using AD groups
  enableStartupSplash: boolean; // Toggle first-load startup splash screen
  enableVerboseStartupStatus: boolean; // Toggle detailed startup stage text in splash
  startupSplashCompletionDelayMs: number; // Delay before hiding splash after ready
  splashVariant: "classic" | "elegant"; // UI style of the startup splash screen
  helpEmail: string; // Support email value configurable from property pane
  helpGuideUrl: string; // Support guide URL configurable from property pane
  enableFavorites: boolean; // Opt-in: show Favorites star + nav link
  enableRecentlyViewed: boolean; // Opt-in: track recently viewed + nav link
  enableTilesView: boolean; // Opt-in: tiles results layout (admin can toggle)
  userPreferencesListName?: string; // Custom list for user prefs (favorites, recently viewed)
  testerEmails?: string; // Comma-separated emails that can cycle roles for testing
  initialDocumentId?: number; // Deep-link: SP list item ID from ?spectraDoc= URL param
}
