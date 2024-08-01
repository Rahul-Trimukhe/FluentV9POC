import * as React from 'react'; 
import { IState } from './IState';


import { WebPartContext } from '@microsoft/sp-webpart-base'; 
// import { useEffect, useState } from 'react';

export enum AppMode {
    SharePoint, SharePointLocal, Teams, TeamsLocal, Office, OfficeLocal, Outlook, OutlookLocal
  }

const Context = React.createContext<IState>({
    isDarkTheme: false,
    _appMode: AppMode.SharePoint,
    hasTeamsContext: false,
    environmentMessage: ""
});

export interface IProvider { 
    children: JSX.Element; 
    context: WebPartContext; 
    isDarkTheme: boolean;
    hasTeamsContext: boolean
   _appMode: AppMode;
   environmentMessage: string;
}

const ContextProvider: React.FC<IProvider> = ({ children, context, isDarkTheme, _appMode, hasTeamsContext, environmentMessage }) => {
    return (
        <Context.Provider value={{
            context, isDarkTheme, _appMode, hasTeamsContext, environmentMessage
        }}>
            {children}
        </Context.Provider>
    );
};

export { ContextProvider, Context };