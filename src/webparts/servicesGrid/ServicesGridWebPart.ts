import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'ServicesGridWebPartStrings';
import ServicesGrid from './components/ServicesGrid';
import { IServicesGridProps, IServiceItem } from './components/IServicesGridProps';

export interface IServicesGridWebPartProps {
  sectionTitle: string;
  services: IServiceItem[];
  columns: string;
}

export default class ServicesGridWebPart extends BaseClientSideWebPart<IServicesGridWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IServicesGridProps> = React.createElement(
      ServicesGrid,
      {
        sectionTitle: this.properties.sectionTitle || 'Recent Launches',
        services: this.properties.services || [],
        columns: this.properties.columns || "4",
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
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
          header: { description: 'Configure the Services Grid' },
          groups: [
            {
              groupName: 'Header Settings',
              groupFields: [
                PropertyPaneTextField('sectionTitle', { label: 'Section Title' })
              ]
            },
            {
              groupName: 'Services Catalog',
              groupFields: [
                PropertyFieldCollectionData('services', {
                  key: 'servicesCollectionData',
                  label: 'Manage Cards',
                  panelHeader: 'Manage Cards Collection',
                  manageBtnLabel: 'Manage Cards',
                  value: this.properties.services,
                  fields: [
                    { 
                      id: 'imageUrl', 
                      title: 'Logo/Image (Link or Upload)', 
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
                    { id: 'title',       title: 'Title',          type: CustomCollectionFieldType.string, required: true },
                    { id: 'description', title: 'Description',    type: CustomCollectionFieldType.string },
                    { id: 'linkText',    title: 'Link Text',      type: CustomCollectionFieldType.string },
                    { id: 'linkUrl',     title: 'Link URL',       type: CustomCollectionFieldType.string },
                    { id: 'cardColor',   title: 'Card Accent Color (Hex)', type: CustomCollectionFieldType.string }
                  ],
                  disabled: false
                })
              ]
            },
            {
              groupName: 'Design Settings',
              groupFields: [
                PropertyPaneChoiceGroup('columns', {
                  label: 'Grid Columns',
                  options: [
                    { key: '3', text: '3 Columns' },
                    { key: '4', text: '4 Columns' },
                    { key: '5', text: '5 Columns' }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
