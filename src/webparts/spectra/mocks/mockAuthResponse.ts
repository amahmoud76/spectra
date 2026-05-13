import { IAuthResponse } from "../interfaces/IAuthResponse";

export const mockViewer: IAuthResponse = {
  roles: ["viewer"],
  assets: [
    "Imbruvica",
    "ABBV-383",
    "Rinvoq",
    "Skyrizi",
    "ABBV-1354",
    "ABBV-113",
    "ABBV-0252",
    "ABBV-1623",
    "ABBV-Eye1",
  ],
  exclusions: [],
  authVerified: true,
};

export const mockContributor: IAuthResponse = {
  roles: ["contributor"],
  assets: [
    "Imbruvica",
    "ABBV-383",
    "Rinvoq",
    "Skyrizi",
    "ABBV-1354",
    "ABBV-113",
    "ABBV-0252",
    "ABBV-1623",
    "ABBV-Eye1",
  ],
  exclusions: [],
  authVerified: true,
};

export const mockAdmin: IAuthResponse = {
  roles: ["admin"],
  assets: [
    "Imbruvica",
    "ABBV-383",
    "Rinvoq",
    "Skyrizi",
    "ABBV-1354",
    "ABBV-113",
    "ABBV-0252",
    "ABBV-1623",
    "ABBV-Eye1",
  ],
  exclusions: [],
  authVerified: true,
};

// Returns the right mock based on MOCK_ROLE setting in config.ts
export const getMockAuthResponse = (
  role: "admin" | "contributor" | "viewer",
): IAuthResponse => {
  switch (role) {
    case "admin":
      return mockAdmin;
    case "contributor":
      return mockContributor;
    default:
      return mockViewer;
  }
};
