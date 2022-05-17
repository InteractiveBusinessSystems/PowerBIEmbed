// Scope of AAD app. Use the below configuration to use all the permissions provided in the AAD app through Azure portal.
// Refer https://aka.ms/PowerBIPermissions for complete list of Power BI scopes
export const scopes: string[] = ["https://analysis.windows.net/powerbi/api/Report.Read.All"];

// Client Id (Application Id) of the AAD app.
export const clientId: string = "4eea776d-9c64-4b9e-973b-d51e3e0e5466";

// Id of the workspace where the report is hosted
export const workspaceId: string = "7af23086-163b-4747-bd1c-977d1830d59b";

// Id of the report to be embedded
export const reportId: string = "349849e5-75db-4ae4-a8cd-e19ab7709daa";

// Id of the report page to be embedded
export const reportsectionId: string = "36c2950023335029ab34";

// Tenant Id
export const tenantId: string = '4ec55493-6b1c-4565-a868-2ae940882c82';

// msal config
export const msalConfig = {
  auth: {
    clientId: clientId,
    // authority: `https://login.microsoft.online.com/common/4ec55493-6b1c-4565-a868-2ae940882c82`,
    // redirectUri: 'https://ibsmtg.sharepoint.com/sites/MaryvilleAcademy-SPFXPowerBIWebPart/_layouts/15/workbench.aspx'
  }
};

export const loginRequest = {
  scopes: scopes
};
