import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  type IPropertyPaneField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'IntranetHubWebPartStrings';
import IntranetHub from './components/IntranetHub';
import {
  IIntranetHubProps,
  ILinkCardItem,
  IDashboardStat,
  IDepartmentItem,
  IGroupedCardItem
} from './components/IIntranetHubProps';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface IIntranetHubWebPartProps {
  /* Quick Links */
  showQuickLinks: boolean;
  quickLinksTitle: string;
  quickLinks: ILinkCardItem[];
  /* Dashboard */
  showDashboard: boolean;
  dashboardTitle: string;
  dashboardStats: IDashboardStat[];
  /* Departments */
  showDepartments: boolean;
  departmentsTitle: string;
  departments: IDepartmentItem[];
  /* Tools */
  showToolsGrid: boolean;
  toolsItems: IGroupedCardItem[];
  /* Spotlight */
  showSpotlight: boolean;
  spotlightName: string;
  spotlightRole: string;
  spotlightDescription: string;
  spotlightImageUrl: string;
  /* Materials */
  showMaterials: boolean;
  materialsItems: IGroupedCardItem[];
}

export default class IntranetHubWebPart extends BaseClientSideWebPart<IIntranetHubWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  /* ── Render ── */

  public render(): void {
    const element: React.ReactElement<IIntranetHubProps> = React.createElement(
      IntranetHub,
      {
        showQuickLinks: this.properties.showQuickLinks,
        quickLinksTitle: this.properties.quickLinksTitle || 'Quick Links',
        quickLinks: this.properties.quickLinks || [],
        showDashboard: this.properties.showDashboard,
        dashboardTitle: this.properties.dashboardTitle || 'Dashboard',
        dashboardStats: this.properties.dashboardStats || [],
        showDepartments: this.properties.showDepartments,
        departmentsTitle: this.properties.departmentsTitle || 'Departments',
        departments: this.properties.departments || [],
        showToolsGrid: this.properties.showToolsGrid,
        toolsItems: this.properties.toolsItems || [],
        showSpotlight: this.properties.showSpotlight,
        spotlightName: this.properties.spotlightName || '',
        spotlightRole: this.properties.spotlightRole || '',
        spotlightDescription: this.properties.spotlightDescription || '',
        spotlightImageUrl: this.properties.spotlightImageUrl || '',
        showMaterials: this.properties.showMaterials,
        materialsItems: this.properties.materialsItems || [],
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /* ── Lifecycle ── */

  protected async onInit(): Promise<void> {
    configureFluentUi();
    await super.onInit();
    this._environmentMessage = await this._getEnvironmentMessage();
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let msg: string = '';
          switch (context.app.host.name) {
            case 'Office': msg = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment; break;
            case 'Outlook': msg = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment; break;
            case 'Teams': case 'TeamsModern': msg = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment; break;
            default: msg = strings.UnknownEnvironment;
          }
          return msg;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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

  /* ── Spotlight image custom field ── */

  /* ── PropertyPane configuration (6 pages) ── */

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        /* ── Page 1: Quick Links ── */
        {
          header: { description: 'Quick Links Section' },
          groups: [
            {
              groupName: 'Visibility',
              groupFields: [
                PropertyPaneToggle('showQuickLinks', { label: 'Show Quick Links', onText: 'Visible', offText: 'Hidden' })
              ]
            },
            {
              groupName: 'Content',
              groupFields: [
                PropertyPaneTextField('quickLinksTitle', { label: 'Section Title' }),
                PropertyFieldCollectionData('quickLinks', {
                  key: 'quickLinksData',
                  label: 'Manage Quick Links',
                  panelHeader: 'Quick Links',
                  manageBtnLabel: 'Manage Links',
                  value: this.properties.quickLinks,
                  fields: [
                    { id: 'iconName', title: 'Fluent Icon Name', type: CustomCollectionFieldType.string, required: false },
                    { id: 'title',    title: 'Title',           type: CustomCollectionFieldType.string, required: true },
                    { id: 'linkUrl',  title: 'Link URL',        type: CustomCollectionFieldType.string, required: false }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        },

        /* ── Page 2: Dashboard Stats ── */
        {
          header: { description: 'Dashboard Stats Section' },
          groups: [
            {
              groupName: 'Visibility',
              groupFields: [
                PropertyPaneToggle('showDashboard', { label: 'Show Dashboard', onText: 'Visible', offText: 'Hidden' })
              ]
            },
            {
              groupName: 'Content',
              groupFields: [
                PropertyPaneTextField('dashboardTitle', { label: 'Section Title' }),
                PropertyFieldCollectionData('dashboardStats', {
                  key: 'dashboardStatsData',
                  label: 'Manage Stats',
                  panelHeader: 'Dashboard Statistics',
                  manageBtnLabel: 'Manage Stats',
                  value: this.properties.dashboardStats,
                  fields: [
                    { id: 'label', title: 'Label (e.g. Active Projects:)', type: CustomCollectionFieldType.string, required: true },
                    { id: 'value', title: 'Value (e.g. 132)',              type: CustomCollectionFieldType.string, required: true }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        },

        /* ── Page 3: Departments ── */
        {
          header: { description: 'Departments Section' },
          groups: [
            {
              groupName: 'Visibility',
              groupFields: [
                PropertyPaneToggle('showDepartments', { label: 'Show Departments', onText: 'Visible', offText: 'Hidden' })
              ]
            },
            {
              groupName: 'Content',
              groupFields: [
                PropertyPaneTextField('departmentsTitle', { label: 'Section Title' }),
                PropertyFieldCollectionData('departments', {
                  key: 'departmentsData',
                  label: 'Manage Departments',
                  panelHeader: 'Departments',
                  manageBtnLabel: 'Manage Departments',
                  value: this.properties.departments,
                  fields: [
                    { id: 'iconName', title: 'Fluent Icon Name',    type: CustomCollectionFieldType.string, required: false },
                    { id: 'title',    title: 'Title',               type: CustomCollectionFieldType.string, required: true },
                    { id: 'linkUrl',  title: 'Link URL',            type: CustomCollectionFieldType.string, required: false },
                    { id: 'bgColor',  title: 'Background Color (Hex)', type: CustomCollectionFieldType.string, required: false }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        },

        /* ── Page 4: Tools & Calculators ── */
        {
          header: { description: 'Tools & Calculators Section' },
          groups: [
            {
              groupName: 'Visibility',
              groupFields: [
                PropertyPaneToggle('showToolsGrid', { label: 'Show Tools Section', onText: 'Visible', offText: 'Hidden' })
              ]
            },
            {
              groupName: 'Content',
              groupFields: [
                PropertyFieldCollectionData('toolsItems', {
                  key: 'toolsItemsData',
                  label: 'Manage Tools',
                  panelHeader: 'Tools & Calculators — items are grouped by Section Title',
                  manageBtnLabel: 'Manage Tools',
                  value: this.properties.toolsItems,
                  fields: [
                    { id: 'sectionTitle', title: 'Section Title (group heading)',  type: CustomCollectionFieldType.string, required: true },
                    { id: 'iconName',     title: 'Fluent Icon Name',              type: CustomCollectionFieldType.string, required: false },
                    { id: 'title',        title: 'Card Title',                    type: CustomCollectionFieldType.string, required: true },
                    { id: 'linkUrl',      title: 'Link URL',                      type: CustomCollectionFieldType.string, required: false }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        },

        /* ── Page 5: Employee Spotlight ── */
        {
          header: { description: 'Employee Spotlight Section' },
          groups: [
            {
              groupName: 'Visibility',
              groupFields: [
                PropertyPaneToggle('showSpotlight', { label: 'Show Spotlight', onText: 'Visible', offText: 'Hidden' })
              ]
            },
            {
              groupName: 'Content',
              groupFields: [
                PropertyPaneTextField('spotlightName', { label: 'Employee Name' }),
                PropertyPaneTextField('spotlightRole', { label: 'Role / Title' }),
                PropertyPaneTextField('spotlightDescription', { label: 'Description', multiline: true, rows: 3 }),
                PropertyFieldCollectionData('spotlightImage', {
                  key: 'spotlightImageData',
                  label: 'Manage Employee Photo',
                  panelHeader: 'Employee Photo (Max 1)',
                  manageBtnLabel: 'Manage Photo',
                  value: this.properties.spotlightImageUrl ? [{ imageUrl: this.properties.spotlightImageUrl }] : [],
                  fields: [
                    {
                      id: 'imageUrl',
                      title: 'Image (Link or Upload)',
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item, itemId, onError) => {
                        return React.createElement("div", { style: { display: "flex", flexDirection: "column", gap: "6px" } }, [
                          React.createElement("input", {
                            key: "text-input",
                            type: "text",
                            placeholder: "Paste Image URL...",
                            value: value && !value.toString().startsWith("data:image") ? value : "",
                            onChange: (e: any) => {
                               onUpdate(field.id, e.target.value);
                               // Sync value back to root prop for backwards compatibility if needed, or just use collection
                               (this.properties as any).spotlightImageUrl = e.target.value;
                            },
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
                                    (this.properties as any).spotlightImageUrl = event.target.result;
                                  }
                                };
                                reader.readAsDataURL(file);
                              }
                            }
                          }),
                          value ? React.createElement("img", { key: "preview", src: value, style: { height: "40px", width: "auto", objectFit: "contain", marginTop: "4px" } }) : null
                        ]);
                      }
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        },

        /* ── Page 6: Materials & Resources ── */
        {
          header: { description: 'Materials, Specs & Product Database Section' },
          groups: [
            {
              groupName: 'Visibility',
              groupFields: [
                PropertyPaneToggle('showMaterials', { label: 'Show Materials Section', onText: 'Visible', offText: 'Hidden' })
              ]
            },
            {
              groupName: 'Content',
              groupFields: [
                PropertyFieldCollectionData('materialsItems', {
                  key: 'materialsItemsData',
                  label: 'Manage Materials',
                  panelHeader: 'Materials & Resources — items are grouped by Section Title',
                  manageBtnLabel: 'Manage Materials',
                  value: this.properties.materialsItems,
                  fields: [
                    { id: 'sectionTitle', title: 'Section Title (group heading)',  type: CustomCollectionFieldType.string, required: true },
                    { id: 'iconName',     title: 'Fluent Icon Name',              type: CustomCollectionFieldType.string, required: false },
                    { id: 'title',        title: 'Card Title',                    type: CustomCollectionFieldType.string, required: true },
                    { id: 'linkUrl',      title: 'Link URL',                      type: CustomCollectionFieldType.string, required: false }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
