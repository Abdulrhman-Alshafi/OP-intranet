import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import {
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ItHelpDeskHeroWebPartStrings';
import ItHelpDeskHero from './components/ItHelpDeskHero';
import type { IItHelpDeskHeroProps } from './components/IItHelpDeskHeroProps';
import { configureFluentUi } from '../../common/configureFluentUi';

import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType,
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface IItHelpDeskHeroWebPartProps {
  title: string;
  titleColor: string;
  description: string;
  backgroundColor: string;
  leftColumnWidthPercent: string;
  buttons: {
    label: string;
    iconName: string;
    url: string;
    openInNewTab: boolean;
  }[];
  tiles: {
    title: string;
    linkText: string;
    linkUrl: string;
    imageUrl: string;
    openInNewTab: boolean;
  }[];
}

export default class ItHelpDeskHeroWebPart extends BaseClientSideWebPart<IItHelpDeskHeroWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IItHelpDeskHeroProps> = React.createElement(
      ItHelpDeskHero,
      {
        title: this.properties.title,
        titleColor: this.properties.titleColor,
        description: this.properties.description,
        backgroundColor: this.properties.backgroundColor,
        leftColumnWidthPercent: this.properties.leftColumnWidthPercent,
        buttons: this.properties.buttons || [],
        tiles: this.properties.tiles || [],
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    configureFluentUi();
    await super.onInit();
    this._environmentMessage = await this._getEnvironmentMessage();
  }

  private async _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      const context = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
      let environmentMessage: string = '';
      switch (context.app.host.name) {
        case 'Office':
          environmentMessage = this.isUnifiedApp() ? 'in the Microsoft 365 app (Windows)' : 'in Office';
          break;
        case 'Outlook':
          environmentMessage = this.isUnifiedApp() ? 'in Outlook (Windows)' : 'in Outlook';
          break;
        case 'Teams':
        case 'TeamsModern':
          environmentMessage = this.isUnifiedApp() ? 'in the Teams app (Windows)' : 'in Teams';
          break;
        default:
          environmentMessage = 'in an unrecognized Microsoft 365 host application';
      }
      return environmentMessage;
    }
    return this.context.isServedFromLocalhost
      ? 'from localhost'
      : 'from SharePoint';
  }

  private isUnifiedApp(): boolean {
    return false;
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
        /*──────────────────────────────────────────────────────
          PAGE 1 — Content
        ──────────────────────────────────────────────────────*/
        {
          header: { description: strings.ContentPageHeader },
          groups: [
            {
              groupName: strings.HeaderGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  multiline: false,
                  placeholder: 'e.g. Welcome to the IT help desk',
                }),
                PropertyPaneTextField('titleColor', {
                  label: strings.TitleColorFieldLabel,
                  multiline: false,
                  placeholder: '#0078d4',
                  description: 'Any valid CSS color: #hex, rgb(), or named color.',
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true,
                  rows: 4,
                  placeholder: 'Supporting text shown below the title.',
                }),
              ],
            },
            {
              groupName: strings.ButtonsGroupName,
              groupFields: [
                PropertyFieldCollectionData('buttons', {
                  key: 'buttonsData',
                  label: strings.ButtonsFieldLabel,
                  panelHeader: strings.ButtonsPanelHeader,
                  manageBtnLabel: strings.ButtonsManageLabel,
                  value: this.properties.buttons,
                  fields: [
                    {
                      id: 'label',
                      title: 'Button Label',
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'e.g. Submit a new ticket',
                    },
                    {
                      id: 'iconName',
                      title: 'Fluent Icon Name',
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'e.g. Mail, Search, Settings',
                    },
                    {
                      id: 'url',
                      title: 'Button URL',
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'https://',
                    },
                    {
                      id: 'openInNewTab',
                      title: 'Open in New Tab',
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                    },
                  ],
                  disabled: false,
                }),
              ],
            },
          ],
        },

        /*──────────────────────────────────────────────────────
          PAGE 2 — Image Tiles
        ──────────────────────────────────────────────────────*/
        {
          header: { description: strings.TilesPageHeader },
          groups: [
            {
              groupName: strings.TilesGroupName,
              groupFields: [
                PropertyFieldCollectionData('tiles', {
                  key: 'tilesData',
                  label: strings.TilesFieldLabel,
                  panelHeader: strings.TilesPanelHeader,
                  manageBtnLabel: strings.TilesManageLabel,
                  value: this.properties.tiles,
                  fields: [
                    {
                      id: 'title',
                      title: 'Tile Title',
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'e.g. Returning to the worksite?',
                    },
                    {
                      id: 'linkText',
                      title: 'Link Label (optional)',
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'e.g. Review latest policies →',
                    },
                    {
                      id: 'linkUrl',
                      title: 'Tile Link URL',
                      type: CustomCollectionFieldType.string,
                      required: false,
                      placeholder: 'https://',
                    },
                    {
                      id: 'openInNewTab',
                      title: 'Open in New Tab',
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                    },
                    {
                      id: 'imageUrl',
                      title: 'Image (URL or Upload)',
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate) => {
                        return React.createElement(
                          'div',
                          { style: { display: 'flex', flexDirection: 'column', gap: '6px', paddingBottom: '4px' } },
                          [
                            React.createElement('input', {
                              key: 'url-input',
                              type: 'text',
                              placeholder: 'Paste image URL…',
                              value: value || '',
                              style: {
                                width: '100%',
                                padding: '4px 6px',
                                fontSize: '12px',
                                border: '1px solid #c8c6c4',
                                borderRadius: '2px',
                                boxSizing: 'border-box',
                              },
                              onChange: (e: React.ChangeEvent<HTMLInputElement>) => {
                                onUpdate(field.id, e.target.value);
                              },
                            }),
                            React.createElement(
                              'span',
                              { key: 'or-label', style: { fontSize: '11px', color: '#605e5c', textAlign: 'center' } },
                              '— or upload a file —'
                            ),
                            React.createElement('input', {
                              key: 'file-input',
                              type: 'file',
                              accept: 'image/*',
                              style: { fontSize: '12px' },
                              onChange: (e: React.ChangeEvent<HTMLInputElement>) => {
                                const file = e.target.files && e.target.files[0];
                                if (!file) return;
                                const reader = new FileReader();
                                reader.onload = (ev) => {
                                  if (ev.target && ev.target.result) {
                                    onUpdate(field.id, ev.target.result as string);
                                  }
                                };
                                reader.readAsDataURL(file);
                              },
                            }),
                            value
                              ? React.createElement('img', {
                                  key: 'preview',
                                  src: value,
                                  alt: 'preview',
                                  style: {
                                    height: '48px',
                                    objectFit: 'cover',
                                    borderRadius: '2px',
                                    border: '1px solid #edebe9',
                                  },
                                })
                              : null,
                          ]
                        );
                      },
                    },
                  ],
                  disabled: false,
                }),
              ],
            },
          ],
        },

        /*──────────────────────────────────────────────────────
          PAGE 3 — Appearance
        ──────────────────────────────────────────────────────*/
        {
          header: { description: strings.AppearancePageHeader },
          groups: [
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneTextField('backgroundColor', {
                  label: strings.BackgroundColorFieldLabel,
                  multiline: false,
                  placeholder: '#ffffff',
                  description: 'Background color for the entire web part.',
                }),
                PropertyPaneChoiceGroup('leftColumnWidthPercent', {
                  label: strings.LeftColumnWidthFieldLabel,
                  options: [
                    {
                      key: '30',
                      text: '30% — Narrow left panel',
                      iconProps: { officeFabricIconFontName: 'AlignLeft' },
                    },
                    {
                      key: '40',
                      text: '40% — Balanced (default)',
                      iconProps: { officeFabricIconFontName: 'ColumnLeftTwoThirds' },
                      checked: true,
                    },
                    {
                      key: '50',
                      text: '50% — Equal halves',
                      iconProps: { officeFabricIconFontName: 'ColumnCenter' },
                    },
                  ],
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
