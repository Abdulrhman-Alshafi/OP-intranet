import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'SwiperWebPartStrings';
import Swiper from './components/Swiper';
import { ISwiperProps, ISlide, IStaticTile } from './components/ISwiperProps';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface ISwiperWebPartProps {
  slides: ISlide[];
  staticTiles: IStaticTile[];
  autoplayDelay: string; 
  containerHeight: string;
  enablePagination: boolean;
  enableNavigation: boolean;
  buttonStyle: 'solid' | 'outline' | 'transparent';
}

export default class SwiperWebPart extends BaseClientSideWebPart<ISwiperWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ISwiperProps> = React.createElement(
      Swiper,
      {
        slides: this.properties.slides || [],
        staticTiles: this.properties.staticTiles || [],
        autoplayDelay: this.properties.autoplayDelay ? parseInt(this.properties.autoplayDelay) : 5000,
        containerHeight: this.properties.containerHeight || '500',
        enablePagination: this.properties.enablePagination !== false,
        enableNavigation: this.properties.enableNavigation !== false,
        buttonStyle: this.properties.buttonStyle || 'solid',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    configureFluentUi();
    await super.onInit();
    this._environmentMessage = await this._getEnvironmentMessage();
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { 
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': 
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': 
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': 
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
          header: { description: 'Configure your Swiper slides and settings.' },
          groups: [
            {
              groupName: 'Slides Data',
              groupFields: [
                PropertyFieldCollectionData('slides', {
                  key: 'slidesCollectionData',
                  label: 'Manage Slides',
                  panelHeader: 'Manage Slides Collection',
                  manageBtnLabel: 'Manage Slides',
                  value: this.properties.slides,
                  fields: [
                    { 
                      id: 'image', 
                      title: 'Image (Link or Upload)', 
                      type: CustomCollectionFieldType.custom, 
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return React.createElement("div", { style: { display: "flex", flexDirection: "column", gap: "6px" } }, [
                          React.createElement("input", {
                            key: "text-input",
                            type: "text",
                            placeholder: "Paste Image URL...",
                            value: value && !value.toString().startsWith("data:image") ? value : "",
                            onChange: (e: any) => onUpdate(field.id, e.target.value),
                            style: { padding: '6px', border: '1px solid rgb(96, 94, 92)', borderRadius: '2px', width: '100%' }
                          }),
                          React.createElement("span", { key: "label", style: { fontSize: "12px", color: "rgb(96, 94, 92)" } }, "— OR Upload File —"),
                          React.createElement("input", {
                            key: "file-input",
                            type: "file",
                            accept: "image/*",
                            onChange: (e: any) => {
                              const file = e.target.files[0];
                              if (file) {
                                const reader = new FileReader();
                                reader.onload = (event: any) => {
                                  if (event.target && event.target.result) {
                                    onUpdate(field.id, event.target.result);
                                  }
                                };
                                reader.readAsDataURL(file);
                              }
                            }
                          }),
                          value ? React.createElement("img", { key: "preview", src: value, style: { height: "40px", width: "auto", objectFit: "contain", marginTop: "4px" } }) : null
                        ]);
                      }
                    },
                    { id: 'headline',    title: 'Headline',    type: CustomCollectionFieldType.string },
                    { id: 'description', title: 'Description', type: CustomCollectionFieldType.string },
                    { id: 'buttonText',  title: 'Button Text', type: CustomCollectionFieldType.string },
                    { id: 'buttonLink',  title: 'Button Link', type: CustomCollectionFieldType.string }
                  ],
                  disabled: false
                })
              ]
            },
            {
              groupName: 'Side Tiles (Right Area)',
              groupFields: [
                PropertyFieldCollectionData('staticTiles', {
                  key: 'staticTilesCollectionData',
                  label: 'Manage Side Tiles',
                  panelHeader: 'Manage Side Tiles (Max 4)',
                  manageBtnLabel: 'Manage Tiles',
                  value: this.properties.staticTiles,
                  fields: [
                    { 
                      id: 'image', 
                      title: 'Image Background (Link or Upload)', 
                      type: CustomCollectionFieldType.custom, 
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return React.createElement("div", { style: { display: "flex", flexDirection: "column", gap: "6px" } }, [
                          React.createElement("input", {
                            key: "text-input",
                            type: "text",
                            placeholder: "Paste Image URL...",
                            value: value && !value.toString().startsWith("data:image") ? value : "",
                            onChange: (e: any) => onUpdate(field.id, e.target.value),
                            style: { padding: '6px', border: '1px solid rgb(96, 94, 92)', borderRadius: '2px', width: '100%' }
                          }),
                          React.createElement("span", { key: "label", style: { fontSize: "12px", color: "rgb(96, 94, 92)" } }, "— OR Upload File —"),
                          React.createElement("input", {
                            key: "file-input",
                            type: "file",
                            accept: "image/*",
                            onChange: (e: any) => {
                              const file = e.target.files[0];
                              if (file) {
                                const reader = new FileReader();
                                reader.onload = (event: any) => {
                                  if (event.target && event.target.result) {
                                    onUpdate(field.id, event.target.result);
                                  }
                                };
                                reader.readAsDataURL(file);
                              }
                            }
                          }),
                          value ? React.createElement("img", { key: "preview", src: value, style: { height: "40px", width: "auto", objectFit: "contain", marginTop: "4px" } }) : null
                        ]);
                      }
                    },
                    { id: 'bgColor', title: 'Background Color (Hex)', type: CustomCollectionFieldType.string, placeholder: '#0078d4' },
                    { id: 'title',   title: 'Tile Title', type: CustomCollectionFieldType.string },
                    { id: 'link',    title: 'Link URL',   type: CustomCollectionFieldType.string }
                  ],
                  disabled: false
                })
              ]
            },
            {
              groupName: 'Swiper Settings',
              groupFields: [
                PropertyPaneTextField('containerHeight', { label: 'Container Height (pixels)', value: '500' }),
                PropertyPaneTextField('autoplayDelay', { label: 'Autoplay Delay (ms) - 0 to disable' }),
                PropertyPaneDropdown('buttonStyle', {
                  label: 'Button Style',
                  options: [
                    { key: 'solid', text: 'Solid (Filled)' },
                    { key: 'outline', text: 'Outline' },
                    { key: 'transparent', text: 'Transparent' }
                  ]
                }),
                PropertyPaneToggle('enablePagination', { 
                  label: 'Enable Pagination (Bullets)',
                  checked: this.properties.enablePagination !== false
                }),
                PropertyPaneToggle('enableNavigation', { 
                  label: 'Enable Navigation (Arrows)',
                  checked: this.properties.enableNavigation !== false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
