import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestUserGroupWebPart.module.scss';
import * as strings from 'TestUserGroupWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library'

export interface ITestUserGroupWebPartProps {
  description: string;
}

export default class TestUserGroupWebPart extends BaseClientSideWebPart<ITestUserGroupWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _getListData(): Promise<any[]> {
    // Ensure environment is SharePoint before making API call
    if (Environment.type !== EnvironmentType.SharePoint && Environment.type !== EnvironmentType.ClassicSharePoint) {
      return Promise.resolve([]);
    }
  
    return this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser/groups`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error(`Error fetching groups: ${response.statusText}`);
        }
        return response.json();
      })
      .then((data) => data.value)
      .catch((error) => {
        console.error('Error fetching SharePoint groups:', error);
        return [];
      });
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((groups) => {
        this._renderList(groups);
      })
      .catch((error) => {
        console.error('Error rendering list:', error);
      });
  }
  
  private _renderList(groups: any[]): void {
    let html: string = '';
    if (groups.length === 0) {
      html = '<p>No groups found.</p>';
    } else {
      groups.forEach((group: any) => {
        html += `
          <ul>
            <li>
              <span class="ms-font-l">${group.Title}</span>
            </li>
          </ul>`;
      });
    }
  
    const container = this.domElement.querySelector('#spListContainer');
    if (container) {
      container.innerHTML = html;
    }
  }
  
  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.testUserGroup} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <h4>Your Groups</h4>
      <div id="spListContainer"/>
    </section>`;

    this._renderListAsync();
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
