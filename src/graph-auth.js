import { createNestablePublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';

import { authConfig } from './authconfig.js';

let pca = null;

async function initialize() {
  if (!pca) {
    pca = await createNestablePublicClientApplication(authConfig);
  }

  return pca;
}

export async function getAccessToken(interactive = false) {
  const client = await initialize();

  const request = {
    scopes: ['User.Read'],
  };

  try {
    const result = await client.acquireTokenSilent(request);

    return result.accessToken;
  } catch (error) {
    if (interactive && error instanceof InteractionRequiredAuthError) {
      const result = await client.acquireTokenPopup(request);

      return result.accessToken;
    }

    throw error;
  }
}

export async function getUserProfile(interactive = false) {
  const accessToken = await getAccessToken(interactive);

  const response = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,department,companyName,mail,userPrincipalName,businessPhones,mobilePhone', {
    headers: {
      Authorization: 'Bearer ' + accessToken,
    },
  });

  if (!response.ok) {
    throw new Error('Graph HTTP ' + response.status);
  }

  return await response.json();
}
