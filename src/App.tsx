
import { MsalProvider } from '@azure/msal-react';
import { IPublicClientApplication } from '@azure/msal-browser';

import ProvideAppContext from './AppContext';
import Calendar from './components/Calendar';
import './css/App.css';
import '@bryntum/calendar/calendar.stockholm.css';
import React from 'react';



type AppProps = {
  pca: IPublicClientApplication
};

export default function App({ pca }: AppProps): React.JSX.Element {
    return (
        <MsalProvider instance={pca}>
            <ProvideAppContext>
                <Calendar />
            </ProvideAppContext>
        </MsalProvider>
    );
}