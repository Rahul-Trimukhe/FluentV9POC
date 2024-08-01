import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AppMode } from "../../FluentUiV9PocWebPart";

export interface IState {
    context?: WebPartContext;
    isDarkTheme: boolean;
   _appMode: AppMode;
   hasTeamsContext: boolean;
   environmentMessage: string;
}