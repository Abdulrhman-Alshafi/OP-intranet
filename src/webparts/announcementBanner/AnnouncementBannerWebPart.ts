import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import AnnouncementBanner from './components/AnnouncementBanner';
import { IAnnouncementBannerProps, IAnnouncement } from './components/IAnnouncementBannerProps';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface IAnnouncementBannerWebPartProps {
  announcements: IAnnouncement[];
  showIcon: boolean;
  showDismiss: boolean;
  stackDirection: 'vertical' | 'horizontal';
}

export default class AnnouncementBannerWebPart extends BaseClientSideWebPart<IAnnouncementBannerWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IAnnouncementBannerProps> = React.createElement(
      AnnouncementBanner,
      {
        announcements: this.properties.announcements || [],
        showIcon: this.properties.showIcon !== false,
        showDismiss: this.properties.showDismiss !== false,
        stackDirection: this.properties.stackDirection || 'vertical',
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
          header: { description: 'Configure your Announcement Banners' },
          groups: [
            {
              groupName: 'Announcements',
              groupFields: [
                PropertyFieldCollectionData('announcements', {
                  key: 'announcementsCollectionData',
                  label: 'Manage Announcements',
                  panelHeader: 'Manage Announcements',
                  manageBtnLabel: 'Manage Announcements',
                  value: this.properties.announcements,
                  fields: [
                    {
                      id: 'message',
                      title: 'Message',
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: 'type',
                      title: 'Type',
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        { key: 'info', text: 'Info (Blue)' },
                        { key: 'warning', text: 'Warning (Amber)' },
                        { key: 'success', text: 'Success (Green)' },
                        { key: 'urgent', text: 'Urgent (Red)' }
                      ],
                      required: true
                    },
                    {
                      id: 'linkText',
                      title: 'Link Text',
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'linkUrl',
                      title: 'Link URL',
                      type: CustomCollectionFieldType.string
                    }
                  ],
                  disabled: false
                })
              ]
            },
            {
              groupName: 'Display Settings',
              groupFields: [
                PropertyPaneDropdown('stackDirection', {
                  label: 'Stack Direction',
                  options: [
                    { key: 'vertical', text: 'Vertical (Stacked)' },
                    { key: 'horizontal', text: 'Horizontal (Side by Side)' }
                  ]
                }),
                PropertyPaneToggle('showIcon', { label: 'Show Icons' }),
                PropertyPaneToggle('showDismiss', { label: 'Allow Dismiss' })
              ]
            }
          ]
        }
      ]
    };
  }
}
