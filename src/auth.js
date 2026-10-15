// Nested App Authentication (NAA): the add-in reuses the Office sign-in, no separate login page.
import { createNestablePublicClientApplication } from "@azure/msal-browser";
import { config } from "./config.js";

let pcaPromise;

function getPca() {
  if (!pcaPromise) {
    pcaPromise = createNestablePublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
    });
  }
  return pcaPromise;
}

export function isAuthSupported() {
  return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
}

export async function getToken() {
  const pca = await getPca();
  const request = { scopes: config.scopes };
  try {
    const res = await pca.acquireTokenSilent(request);
    return res.accessToken;
  } catch {
    // First run or consent needed: Office shows a small sign-in/consent popup.
    const res = await pca.acquireTokenPopup(request);
    return res.accessToken;
  }
}
