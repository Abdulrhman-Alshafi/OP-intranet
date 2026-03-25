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
import { RecognitionService } from '../../services/RecognitionService';

import * as strings from 'RecognitionWallWebPartStrings';
import RecognitionWall from './components/RecognitionWall';
import { IRecognitionWallProps } from './components/IRecognitionWallProps';

export interface IRecognitionWallWebPartProps {
  title: string;
  refreshInterval: number;
  showEmployeeOfMonth: boolean;
  postsPerPage: number;
}

export default class RecognitionWallWebPart extends BaseClientSideWebPart<IRecognitionWallWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _currentUserId: number = 0;
  private _isAdmin: boolean = false;

  public async render(): Promise<void> {
    // Resolve current user ID if not yet known
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
        console.warn('RecognitionWallWebPart: Could not resolve current user', error);
      }
    }

    const element: React.ReactElement<IRecognitionWallProps> = React.createElement(
      RecognitionWall,
      {
        title: this.properties.title || 'Recognition Wall',
        refreshInterval: this.properties.refreshInterval || 5,
        showEmployeeOfMonth: this.properties.showEmployeeOfMonth !== false,
        postsPerPage: this.properties.postsPerPage || 10,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        currentUserId: this._currentUserId,
        currentUserEmail: this.context.pageContext.user.email || '',
        currentUserName: this.context.pageContext.user.displayName || '',
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        isAdmin: this._isAdmin
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
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
                  min: 1,
                  max: 60,
                  step: 1,
                  showValue: true,
                  value: this.properties.refreshInterval || 5
                }),
                PropertyPaneToggle('showEmployeeOfMonth', {
                  label: strings.ShowEmployeeOfMonthLabel,
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneSlider('postsPerPage', {
                  label: strings.PostsPerPageLabel,
                  min: 5,
                  max: 50,
                  step: 5,
                  showValue: true,
                  value: this.properties.postsPerPage || 10
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
