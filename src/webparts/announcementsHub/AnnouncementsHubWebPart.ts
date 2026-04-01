import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import * as strings from 'AnnouncementsHubWebPartStrings';
import AnnouncementsHub from './components/AnnouncementsHub';
import { IAnnouncementsHubProps } from './components/IAnnouncementsHubProps';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface IAnnouncementsHubWebPartProps {
  title: string;
  layoutMode: string;
  itemsToDisplay: number;
  sortOrder: string;
  showImages: boolean;
  enableAnimations: boolean;
  enableCategoryColors: boolean;
  accentColor: string;
}

export default class AnnouncementsHubWebPart extends BaseClientSideWebPart<IAnnouncementsHubWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _currentUserId: number = 0;
  private _isAdmin: boolean = false;

  public async render(): Promise<void> {
    // Resolve current user ID and admin status
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
        console.warn('AnnouncementsHubWebPart: Could not resolve current user', error);
      }
    }

    const element: React.ReactElement<IAnnouncementsHubProps> = React.createElement(
      AnnouncementsHub,
      {
        title: this.properties.title || 'OP Announcements Hub',
        layoutMode: (this.properties.layoutMode as 'grid' | 'list') || 'grid',
        itemsToDisplay: this.properties.itemsToDisplay || 9,
        sortOrder: (this.properties.sortOrder as 'newest' | 'oldest') || 'newest',
        showImages: this.properties.showImages !== false,
        enableAnimations: this.properties.enableAnimations !== false,
        enableCategoryColors: this.properties.enableCategoryColors !== false,
        accentColor: this.properties.accentColor || '#0078d4',
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        currentUserId: this._currentUserId,
        isAdmin: this._isAdmin,
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
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneDropdown('layoutMode', {
                  label: strings.LayoutModeLabel,
                  options: [
                    { key: 'grid', text: 'Card Grid View' },
                    { key: 'list', text: 'Compact List View' }
                  ],
                  selectedKey: this.properties.layoutMode || 'grid'
                }),
                PropertyPaneSlider('itemsToDisplay', {
                  label: strings.ItemsToDisplayLabel,
                  min: 1,
                  max: 20,
                  step: 1,
                  showValue: true,
                  value: this.properties.itemsToDisplay || 9
                }),
                PropertyPaneDropdown('sortOrder', {
                  label: strings.SortOrderLabel,
                  options: [
                    { key: 'newest', text: 'Newest First' },
                    { key: 'oldest', text: 'Oldest First' }
                  ],
                  selectedKey: this.properties.sortOrder || 'newest'
                }),
                PropertyPaneToggle('showImages', {
                  label: strings.ShowImagesLabel,
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('enableAnimations', {
                  label: strings.EnableAnimationsLabel,
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            },
            {
              groupName: strings.ThemeGroupName,
              groupFields: [
                PropertyPaneToggle('enableCategoryColors', {
                  label: strings.EnableCategoryColorsLabel,
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyFieldColorPicker('accentColor', {
                  label: strings.AccentColorLabel,
                  selectedColor: this.properties.accentColor || '#0078d4',
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'accentColorField'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
