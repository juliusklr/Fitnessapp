import { PublicClientApplication } from '@azure/msal-browser';

const msalConfig = {
  auth: {
    clientId: '812777cb-5b2f-4d0b-9e23-5708c6df6948',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'localStorage',
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ['Files.ReadWrite', 'User.Read'],
};

export const msalInstance = new PublicClientApplication(msalConfig);
