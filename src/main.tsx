import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';

import {
    PublicClientApplication,
    EventType,
    EventMessage,
    AuthenticationResult
} from '@azure/msal-browser';

import config from './Config';
import App from './App';
import './css/index.css';


const msalInstance = new PublicClientApplication({
    auth : {
        clientId    : config.appId,
        redirectUri : config.redirectUri
    },
    cache : {
        cacheLocation          : 'sessionStorage',
        storeAuthStateInCookie : true
    }
});

// Check if there are already accounts in the browser session
// If so, set the first account as the active account
const accounts = msalInstance.getAllAccounts();
if (accounts && accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
}

msalInstance.addEventCallback((event: EventMessage) => {
    if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
    // Set the active account - this simplifies token acquisition
        const authResult = event.payload as AuthenticationResult;
        msalInstance.setActiveAccount(authResult.account);
    }
});

createRoot(document.getElementById('root') as HTMLElement).render(
    <StrictMode>
        <App pca={msalInstance} />
    </StrictMode>
);
