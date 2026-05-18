import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { initializeIcons } from "@fluentui/react/lib/Icons";
import { SPECTRA } from "./components/SPECTRA";
import { IWebPartProps } from "./interfaces/IWebPartProps";
import {
  PAGE_SIZE_DEFAULT,
  PAGE_SIZE_OPTIONS,
  SP_DOCUMENT_LIBRARY,
  INACTIVITY_TIMEOUT_DEFAULT_MINUTES,
  INACTIVITY_TIMEOUT_OPTIONS_MINUTES,
  USE_MOCK,
  USE_AD_GROUPS,
  MOCK_ROLE,
  HELP_EMAIL,
  HELP_GUIDE_URL,
} from "./config/config";

export interface IPeaksWebPartProps {
  title: string;
  pageSize: number;
  inactivityTimeoutMinutes: number;
  documentLibrary: string;
  useMock: boolean;
  useAdGroups: boolean;
  mockRole: "admin" | "contributor" | "viewer";
  enableStartupSplash: boolean;
  enableVerboseStartupStatus: boolean;
  startupSplashCompletionDelayMs: number;
  helpEmail: string;
  helpGuideUrl: string;
  enableEnhancedTableStyle: boolean;
}

export default class PeaksWebPart extends BaseClientSideWebPart<IPeaksWebPartProps> {
  protected onInit(): Promise<void> {
    initializeIcons();
    return super.onInit();
  }

  public render(): void {
    const urlParams = new URLSearchParams(window.location.search);
    const isSpectraDev = urlParams.get("spectraDev") === "1";

    const props: IWebPartProps = {
      context: this.context,
      pageSize: this.properties.pageSize || PAGE_SIZE_DEFAULT,
      title: this.properties.title || "SPECTRA Document Repository",
      userEmail: this.context.pageContext.user.email,
      enableDevRoleSwitch: isSpectraDev,
      documentLibrary: this.properties.documentLibrary || SP_DOCUMENT_LIBRARY,
      inactivityTimeoutMinutes:
        this.properties.inactivityTimeoutMinutes ||
        INACTIVITY_TIMEOUT_DEFAULT_MINUTES,
      useMock: this.properties.useMock ?? USE_MOCK,
      useAdGroups: this.properties.useAdGroups ?? USE_AD_GROUPS,
      mockRole: this.properties.mockRole || MOCK_ROLE,
      enableStartupSplash: this.properties.enableStartupSplash ?? true,
      enableVerboseStartupStatus:
        this.properties.enableVerboseStartupStatus ?? true,
      startupSplashCompletionDelayMs:
        this.properties.startupSplashCompletionDelayMs ?? 700,
      helpEmail: this.properties.helpEmail || HELP_EMAIL,
      helpGuideUrl: this.properties.helpGuideUrl || HELP_GUIDE_URL,
      enableEnhancedTableStyle: (this.properties.enableEnhancedTableStyle ?? false) || isSpectraDev,
      isDevMode: isSpectraDev,
    };

    ReactDom.render(React.createElement(SPECTRA, props), this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configure the SPECTRA Document Repository web part.",
          },
          groups: [
            {
              groupName: "Display Settings",
              groupFields: [
                PropertyPaneTextField("title", {
                  label: "Web part title",
                  value: "SPECTRA Document Repository",
                }),
                PropertyPaneDropdown("pageSize", {
                  label: "Documents per page",
                  selectedKey: PAGE_SIZE_DEFAULT,
                  options: PAGE_SIZE_OPTIONS.map((size) => ({
                    key: size,
                    text:
                      size === PAGE_SIZE_DEFAULT
                        ? `${size} (default)`
                        : String(size),
                  })),
                }),
                PropertyPaneDropdown("inactivityTimeoutMinutes", {
                  label: "Inactivity timeout (minutes)",
                  selectedKey: INACTIVITY_TIMEOUT_DEFAULT_MINUTES,
                  options: INACTIVITY_TIMEOUT_OPTIONS_MINUTES.map(
                    (minutes) => ({
                      key: minutes,
                      text:
                        minutes === INACTIVITY_TIMEOUT_DEFAULT_MINUTES
                          ? `${minutes} (default)`
                          : String(minutes),
                    }),
                  ),
                }),
                PropertyPaneTextField("documentLibrary", {
                  label: "Document library name",
                  value: SP_DOCUMENT_LIBRARY,
                }),
                PropertyPaneToggle("useMock", {
                  label: "Use Demo Mode (Mock Data)",
                  onText: "On",
                  offText: "Off",
                  checked: USE_MOCK,
                }),
                PropertyPaneToggle("useAdGroups", {
                  label: "Use AD Group Membership",
                  onText: "On",
                  offText: "Off",
                  checked: USE_AD_GROUPS,
                }),
                PropertyPaneDropdown("mockRole", {
                  label: "Mock role (used when AD groups are off)",
                  selectedKey: MOCK_ROLE,
                  options: [
                    { key: "admin", text: "Administrator" },
                    { key: "contributor", text: "Contributor" },
                    { key: "viewer", text: "View Only" },
                  ],
                }),
                PropertyPaneToggle("enableStartupSplash", {
                  label: "Show Startup Splash Screen",
                  onText: "On",
                  offText: "Off",
                  checked: true,
                }),
                PropertyPaneToggle("enableVerboseStartupStatus", {
                  label: "Show Detailed Startup Status",
                  onText: "On",
                  offText: "Off",
                  checked: true,
                }),
                PropertyPaneDropdown("startupSplashCompletionDelayMs", {
                  label: "Startup splash completion delay (ms)",
                  selectedKey: 700,
                  options: [
                    { key: 300, text: "300" },
                    { key: 500, text: "500" },
                    { key: 700, text: "700 (default)" },
                    { key: 1000, text: "1000" },
                    { key: 1500, text: "1500" },
                  ],
                }),
                PropertyPaneTextField("helpEmail", {
                  label: "Support email",
                  value: HELP_EMAIL,
                }),
                PropertyPaneTextField("helpGuideUrl", {
                  label: "Help guide URL",
                  value: HELP_GUIDE_URL,
                }),
                PropertyPaneToggle("enableEnhancedTableStyle", {
                  label: "Enhanced table style (admins only)",
                  onText: "On",
                  offText: "Off",
                  checked: false,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
