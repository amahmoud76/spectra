import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

export interface IPeopleSuggestion {
  displayName: string;
  email: string;
}

// Mock fallback — matches unique createdBy values in mock documents
const MOCK_PEOPLE: IPeopleSuggestion[] = [
  { displayName: "Anderson, Karen",  email: "karen.anderson@abbvie.com"  },
  { displayName: "Chen, Lisa",       email: "lisa.chen@abbvie.com"       },
  { displayName: "Davis, Emily",     email: "emily.davis@abbvie.com"     },
  { displayName: "Johnson, Sarah",   email: "sarah.johnson@abbvie.com"   },
  { displayName: "Smith, John",      email: "john.smith@abbvie.com"      },
  { displayName: "Taylor, Robert",   email: "robert.taylor@abbvie.com"   },
  { displayName: "Williams, Mark",   email: "mark.williams@abbvie.com"   },
];

export class PeopleService {
  private context: WebPartContext;
  private useMock: boolean;

  constructor(context: WebPartContext, useMock: boolean = false) {
    this.context = context;
    this.useMock = useMock;
  }

  async searchUsers(query: string): Promise<IPeopleSuggestion[]> {
    if (this.useMock) {
      const q = query.toLowerCase();
      return MOCK_PEOPLE.filter(
        (p) =>
          p.displayName.toLowerCase().includes(q) ||
          p.email.toLowerCase().includes(q),
      );
    }

    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const url =
        `${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface` +
        `.clientPeoplePickerSearchUser`;

      const resp = await this.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
          },
          body: JSON.stringify({
            queryParams: {
              AllowEmailAddresses: true,
              AllowMultipleEntities: false,
              AllUrlZones: false,
              MaximumEntitySuggestions: 15,
              PrincipalSource: 15, // AD + SP users
              PrincipalType: 1,    // User only (not groups)
              QueryString: query,
            },
          }),
        },
      );

      if (!resp.ok) return [];

      const data = await resp.json();
      // The API wraps its result as a JSON-encoded string inside `.value`
      const raw: string =
        typeof data.value === "string"
          ? data.value
          : JSON.stringify(data.value ?? []);
      const users: Record<string, unknown>[] = JSON.parse(raw);

      return users
        .filter((u) => u.EntityType === "User")
        .map((u) => {
          const entityData = (u.EntityData as Record<string, string>) ?? {};
          return {
            displayName:
              (u.DisplayText as string) || entityData.DisplayName || "",
            email:
              entityData.Email ||
              entityData.AccountName ||
              (u.Key as string) ||
              "",
          };
        })
        .filter((p) => p.displayName.length > 0);
    } catch {
      return [];
    }
  }
}
