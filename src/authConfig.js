// ─────────────────────────────────────────────────────────────────────────────
// STEP 1: Fill in your Azure AD app registration details here.
// See AZURE_SETUP.md for step-by-step instructions.
// ─────────────────────────────────────────────────────────────────────────────

export const msalConfig = {
    auth: {
          clientId: "a575587c-9868-4fd7-8268-1556ac6308fb",
          authority: "https://login.microsoftonline.com/0812948c-d8a4-4cd0-914e-59942e064343",
          redirectUri: window.location.origin,  // Must match a Redirect URI registered in Azure
    },
    cache: {
          cacheLocation: "sessionStorage",
          storeAuthStateInCookie: false,
    },
};

// Scopes needed to read/write to SharePoint and send email
export const loginRequest = {
    scopes: ["Sites.ReadWrite.All", "Mail.Send", "User.Read", "People.Read"],
};

// ─────────────────────────────────────────────────────────────────────────────
// STEP 2: Confirm your SharePoint site URL and list name.
// ─────────────────────────────────────────────────────────────────────────────

export const sharepointConfig = {
    siteUrl: "https://allstardriver.sharepoint.com/sites/ServiceExcellenceDepartment-ALL-CustomerServiceTeam",
    listName: "Support Quality Assurance",
    evaluatorAgentsListName: "QA_EvaluatorAgents",
    assignmentsListName: "QA_Assignments",
};
