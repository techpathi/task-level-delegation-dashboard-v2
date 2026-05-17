import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'TaskLevelDelegationWebPartStrings';
import TaskLevelDelegation from './components/TaskLevelDelegation';
import { ITaskLevelDelegationProps } from './components/ITaskLevelDelegationProps';

export interface ITaskLevelDelegationWebPartProps {
  description: string;
  nintexApiBaseUrl: string;
  tokenListUrl: string;
  tokenColumnName: string;
  tokenTitleValue: string;
  paginationSize: number;
}

export default class TaskLevelDelegationWebPart extends BaseClientSideWebPart<ITaskLevelDelegationWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ITaskLevelDelegationProps> = React.createElement(
      TaskLevelDelegation,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        nintexApiBaseUrl: this.properties.nintexApiBaseUrl,
        tokenListUrl: this.properties.tokenListUrl,
        tokenColumnName: this.properties.tokenColumnName,
        tokenTitleValue: this.properties.tokenTitleValue,
        paginationSize: this.properties.paginationSize || 50
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
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
                }),
                PropertyPaneTextField('nintexApiBaseUrl', {
                  label: "Nintex API Base URL"
                }),
                PropertyPaneTextField('tokenListUrl', {
                  label: "Token List URL"
                }),
                PropertyPaneTextField('tokenColumnName', {
                  label: "Token Column Name"
                }),
                PropertyPaneTextField('tokenTitleValue', {
                  label: "Token Title Filter Value"
                }),
                PropertyPaneSlider('paginationSize', {
                  label: 'Pagination Size',
                  min: 5,
                  max: 500,
                  step: 5,
                  value: 50
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
