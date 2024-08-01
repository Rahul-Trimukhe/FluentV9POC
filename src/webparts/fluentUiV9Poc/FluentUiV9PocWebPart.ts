import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FluentUiV9PocWebPartStrings';
import FluentUiV9Poc from './components/FluentUiV9Poc';
import { IFluentUiV9PocProps } from './components/IFluentUiV9PocProps';  
import { ContextProvider } from './components/context';

export interface IFluentUiV9PocWebPartProps {
  description: string;

}

export enum AppMode {
  SharePoint, SharePointLocal, Teams, TeamsLocal, Office, OfficeLocal, Outlook, OutlookLocal
}

export default class FluentUiV9PocWebPart extends BaseClientSideWebPart<IFluentUiV9PocWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _appMode: AppMode = AppMode.SharePoint;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IFluentUiV9PocProps> = React.createElement(
      FluentUiV9Poc,
      {
      }
    ); 

    const cntextProvider = React.createElement(ContextProvider, {
      children: element,
      context: this.context, 
      isDarkTheme: this._isDarkTheme,
      _appMode: this._appMode,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
    });

    ReactDom.render(cntextProvider, this.domElement);
  }

  protected async onInit(): Promise<void> {
    const _l = this.context.isServedFromLocalhost;
    if(this.context.sdks.microsoftTeams) {
      const teamsCtx = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      switch(teamsCtx.app.host.name.toLowerCase()) {
        case 'teams': _l ? AppMode.TeamsLocal : AppMode.Teams;
        case 'outlook': _l ? AppMode.OutlookLocal : AppMode.Outlook;
        case 'office': _l ? AppMode.Office : AppMode.OfficeLocal;
      }
    } else this._appMode = _l ? AppMode.SharePointLocal : AppMode.SharePoint;

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
