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
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'CvRecommendationWebPartStrings';
import CvRecommendation from './components/CvRecommendation';
import { ICvRecommendationProps } from './components/ICvRecommendationProps';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface ICvRecommendationWebPartProps {
  title: string;
  accentColor: string;
  itemsPerPage: number;
}

export default class CvRecommendationWebPart extends BaseClientSideWebPart<ICvRecommendationWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _currentUserId: number = 0;
  private _currentUserName: string = '';

  public async render(): Promise<void> {
    if (this._currentUserId === 0) {
      try {
        const res = await this.context.spHttpClient.get(
          `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser?$select=Id,Title`,
          SPHttpClient.configurations.v1
        );
        if (res.ok) {
          const json = await res.json();
          this._currentUserId = json.Id;
          this._currentUserName = json.Title ?? '';
        }
      } catch (e) {
        console.warn('CvRecommendationWebPart: Could not resolve current user', e);
      }
    }

    const element: React.ReactElement<ICvRecommendationProps> = React.createElement(
      CvRecommendation,
      {
        title: this.properties.title || 'CV Recommendation',
        accentColor: this.properties.accentColor || '#0078d4',
        itemsPerPage: this.properties.itemsPerPage || 10,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        currentUserId: this._currentUserId,
        currentUserName: this._currentUserName,
        isDarkTheme: this._isDarkTheme
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    configureFluentUi();
    return super.onInit();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--bodyBackground', semanticColors.bodyBackground || null);
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
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneSlider('itemsPerPage', {
                  label: strings.ItemsPerPageLabel,
                  min: 5,
                  max: 50,
                  step: 5
                }),
                PropertyPaneTextField('accentColor', {
                  label: strings.AccentColorLabel,
                  placeholder: '#0078d4'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
