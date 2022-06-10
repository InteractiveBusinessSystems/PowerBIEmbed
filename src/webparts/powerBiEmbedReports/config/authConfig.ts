// Scope of AAD app. Use the below configuration to use all the permissions provided in the AAD app through Azure portal.

// Refer https://aka.ms/PowerBIPermissions for complete list of Power BI scopes
export const scopes: string[] = [
  "https://analysis.windows.net/powerbi/api/Report.Read.All",
  "https://analysis.windows.net/powerbi/api/Workspace.Read.All",
  "https://analysis.windows.net/powerbi/api/Dataset.Read.All",
];

// Client Id (Application Id) of the AAD app.
export const clientId: string = "431d5587-d0e9-4698-9a8f-5297bf4e3412";

// Id of the workspace where the report is hosted
export const workspaceId: string = "7af23086-163b-4747-bd1c-977d1830d59b";

// Id of the report to be embedded
export const reportId: string = "349849e5-75db-4ae4-a8cd-e19ab7709daa";

// Id of the report page to be embedded
export const reportsectionId: string = "36c2950023335029ab34";

// Tenant Id
export const tenantId: string = '4ec55493-6b1c-4565-a868-2ae940882c82';

// Authority Url
export const authorityUrl: string = `https://login.microsoftonline.com/common/v2.0`;

// Resource
export const resource: string = 'https://analysis.windows.net/powerbi/api';

// Master User username
export const MUUserName: string = 'sdarroch@ibs365.com';

// Master User password
export const MUpassword: string = 'B@c0n1sMyF@v';

const cacheLocation: 'sessionStorage' = 'sessionStorage';

// msal config
export const msalConfig = {
  auth: {
    clientId: clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    knownAuthorities: [`https://login.microsoftonline.com/${tenantId}`]
    // redirectUri: 'https://127.0.0.1'
  }
};

export const loginRequest = {
  scopes: scopes
};
