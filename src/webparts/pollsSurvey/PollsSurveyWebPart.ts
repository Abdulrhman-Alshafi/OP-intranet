import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'PollsSurveyWebPartStrings';
import PollsSurvey from './components/PollsSurvey';
import { IPollsSurveyProps } from './components/IPollsSurveyProps';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface IPollsSurveyWebPartProps {
  title: string;
  refreshInterval: number;
  allowAnonymous: boolean;
  pollsPerPage: number;
  autoHideDays: number;
}

export default class PollsSurveyWebPart extends BaseClientSideWebPart<IPollsSurveyWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _currentUserId: number = 0;
  private _isAdmin: boolean = false;

  public async render(): Promise<void> {
    // Resolve current user's SharePoint numeric ID + site admin status
    if (this._currentUserId === 0) {
      try {
        const response = await this.context.spHttpClient.get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser?$select=Id,IsSiteAdmin`,
          SPHttpClient.configurations.v1
        );
        if (response.ok) {
          const json = await response.json();
          this._currentUserId = json.Id;
          this._isAdmin = !!json.IsSiteAdmin;
        }
      } catch (error) {
        console.warn('PollsSurveyWebPart: Could not resolve current user', error);
      }
    }

    const element: React.ReactElement<IPollsSurveyProps> = React.createElement(
      PollsSurvey,
      {
        title: this.properties.title || 'Polls & Quick Surveys',
        refreshInterval: this.properties.refreshInterval || 30,
        allowAnonymous: this.properties.allowAnonymous || false,
        pollsPerPage: this.properties.pollsPerPage || 5,
        isAdmin: this._isAdmin,
        autoHideDays: this.properties.autoHideDays || 30,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        currentUserId: this._currentUserId,
        currentUserDisplayName: this.context.pageContext.user.displayName || '',
        currentUserEmail: this.context.pageContext.user.email || '',
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    configureFluentUi();
    await super.onInit();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
      this.domElement.style.setProperty('--neutralPrimary', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--neutralSecondary', semanticColors.bodySubtext || null);
      this.domElement.style.setProperty('--neutralLight', semanticColors.bodyFrameDivider || null);
      this.domElement.style.setProperty('--neutralLighter', semanticColors.bodyBackground || null);
      this.domElement.style.setProperty('--white', semanticColors.bodyBackground || null);
      this.domElement.style.setProperty('--themePrimary', semanticColors.link || null);
      this.domElement.style.setProperty('--themeDark', semanticColors.linkHovered || null);
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneSlider('refreshInterval', {
                  label: strings.RefreshIntervalLabel,
                  min: 10,
                  max: 300,
                  step: 10,
                  showValue: true,
                  value: this.properties.refreshInterval || 30
                }),
                PropertyPaneToggle('allowAnonymous', {
                  label: strings.AllowAnonymousLabel,
                  onText: 'Yes',
                  offText: 'No'
                }),
                PropertyPaneSlider('pollsPerPage', {
                  label: strings.PollsPerPageLabel,
                  min: 1,
                  max: 20,
                  step: 1,
                  showValue: true,
                  value: this.properties.pollsPerPage || 5
                }),
                PropertyPaneSlider('autoHideDays', {
                  label: strings.AutoHideDaysLabel,
                  min: 1,
                  max: 365,
                  step: 1,
                  showValue: true,
                  value: this.properties.autoHideDays || 30
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
