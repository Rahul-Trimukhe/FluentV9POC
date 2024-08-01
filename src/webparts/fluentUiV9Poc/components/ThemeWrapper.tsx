import { FluentProvider, IdPrefixProvider, webLightTheme, Theme, webDarkTheme, teamsDarkTheme, teamsLightTheme } from "@fluentui/react-components";
import * as React from "react";
import { AppMode } from "../FluentUiV9PocWebPart";
import { Context } from "./context";

const ThemeWrapper: React.FC<{}> = ({ children }) => {
    const { _appMode, isDarkTheme } = React.useContext(Context);
    const getCurrentTheme = () : Theme => {
        return (_appMode === AppMode.Teams || _appMode === AppMode.TeamsLocal) ? (isDarkTheme ? teamsDarkTheme : teamsLightTheme) : ((_appMode === AppMode.SharePoint || _appMode === AppMode.SharePointLocal) ? (isDarkTheme ? webDarkTheme : webLightTheme) :(isDarkTheme ? webDarkTheme : webLightTheme));
    }
    return (
        <>
            <IdPrefixProvider value="APP1-">
                <FluentProvider theme={getCurrentTheme()}>{children}</FluentProvider>
            </IdPrefixProvider>
        </>
    )
}

export default ThemeWrapper;