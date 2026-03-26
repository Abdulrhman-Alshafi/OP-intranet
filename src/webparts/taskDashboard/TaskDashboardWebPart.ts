import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from 'TaskDashboardWebPartStrings';
import TaskDashboard from './components/TaskDashboard';
import { ITaskDashboardProps, ISelectedPlan } from './components/ITaskDashboardProps';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface ITaskDashboardWebPartProps {
  title: string;
  enablePlanner: boolean;
  enableSharePoint: boolean;
  sharePointListId: string;
  selectedPlans: ISelectedPlan[];
  groupBy: 'status' | 'source' | 'dueDate' | 'none';
  sortBy: 'dueDate' | 'priority' | 'title';
  showCompleted: boolean;
  highlightOverdue: boolean;
  refreshInterval: number;
  overdueColor: string;
  inProgressColor: string;
  completedColor: string;
  showProgressBars: boolean;
  showSourceBadges: boolean;
}

export default class TaskDashboardWebPart extends BaseClientSideWebPart<ITaskDashboardWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _graphClient: MSGraphClientV3 | undefined;

  public async render(): Promise<void> {
    // Initialize Graph client
    if (!this._graphClient) {
      try {
        this._graphClient = await this.context.msGraphClientFactory.getClient('3');
      } catch (error) {
        console.error('Failed to initialize Graph client:', error);
        this._renderError('Failed to initialize Microsoft Graph. Please ensure proper permissions are granted.');
        return;
      }
    }

    const element: React.ReactElement<ITaskDashboardProps> = React.createElement(
      TaskDashboard,
      {
        title: this.properties.title || 'My Tasks',
        enablePlanner: this.properties.enablePlanner !== false,
        enableSharePoint: this.properties.enableSharePoint || false,
        sharePointListId: this.properties.sharePointListId || '',
        selectedPlans: this.properties.selectedPlans || [],
        groupBy: this.properties.groupBy || 'status',
        sortBy: this.properties.sortBy || 'dueDate',
        showCompleted: this.properties.showCompleted !== false,
        highlightOverdue: this.properties.highlightOverdue !== false,
        refreshInterval: this.properties.refreshInterval || 5,
        overdueColor: this.properties.overdueColor || '#d13438',
        inProgressColor: this.properties.inProgressColor || '#0078d4',
        completedColor: this.properties.completedColor || '#107c10',
        showProgressBars: this.properties.showProgressBars !== false,
        showSourceBadges: this.properties.showSourceBadges !== false,
        graphClient: this._graphClient,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _renderError(message: string): void {
    this.domElement.innerHTML = `
      <div style="padding: 20px; background: #fef0f1; border: 1px solid #d13438; border-radius: 4px; color: #323130;">
        <h3 style="margin: 0 0 10px 0; color: #d13438;">Configuration Error</h3>
        <p style="margin: 0;">${message}</p>
      </div>
    `;
  }

  protected async onInit(): Promise<void> {
    configureFluentUi();
    await super.onInit();
    this._environmentMessage = await this._getEnvironmentMessage();
    return Promise.resolve();
  }

  private async _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      try {
        const context = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
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
      } catch (error) {
        return strings.UnknownEnvironment;
      }
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

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
              groupName: strings.DataSourcesGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneToggle('enablePlanner', {
                  label: strings.EnablePlannerLabel,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('enableSharePoint', {
                  label: strings.EnableSharePointLabel,
                  onText: 'Enabled',
                  offText: 'Disabled'
                })
              ]
            },
            {
              groupName: 'SharePoint Configuration',
              groupFields: [
                PropertyFieldListPicker('sharePointListId', {
                  label: strings.SharePointListLabel,
                  selectedList: this.properties.sharePointListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: !this.properties.enableSharePoint,
                  baseTemplate: 171, // Tasks list template
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                })
              ]
            },
            {
              groupName: strings.PlanSelectionGroupName,
              groupFields: [
                PropertyFieldCollectionData('selectedPlans', {
                  key: 'selectedPlansCollection',
                  label: strings.SelectPlansLabel,
                  panelHeader: 'Select Planner Plans',
                  manageBtnLabel: 'Select Plans',
                  value: this.properties.selectedPlans,
                  fields: [
                    {
                      id: 'id',
                      title: 'Plan ID',
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Enter Plan ID'
                    },
                    {
                      id: 'title',
                      title: 'Plan Title',
                      type: CustomCollectionFieldType.string,
                      required: true,
                      placeholder: 'Enter Plan Title'
                    }
                  ],
                  disabled: !this.properties.enablePlanner,
                  panelDescription: 'Add specific Planner plans to monitor. Leave empty to show all plans you have access to.'
                })
              ]
            },
            {
              groupName: strings.DisplaySettingsGroupName,
              groupFields: [
                PropertyPaneDropdown('groupBy', {
                  label: strings.GroupByLabel,
                  options: [
                    { key: 'status', text: 'Status' },
                    { key: 'source', text: 'Source (Planner, SharePoint)' },
                    { key: 'dueDate', text: 'Due Date (Overdue, Today, This Week, Later)' },
                    { key: 'none', text: 'No Grouping' }
                  ]
                }),
                PropertyPaneDropdown('sortBy', {
                  label: strings.SortByLabel,
                  options: [
                    { key: 'dueDate', text: 'Due Date' },
                    { key: 'priority', text: 'Priority' },
                    { key: 'title', text: 'Title (A-Z)' }
                  ]
                }),
                PropertyPaneToggle('showCompleted', {
                  label: strings.ShowCompletedLabel,
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('highlightOverdue', {
                  label: strings.HighlightOverdueLabel,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneSlider('refreshInterval', {
                  label: strings.RefreshIntervalLabel,
                  min: 1,
                  max: 60,
                  step: 1,
                  showValue: true,
                  value: this.properties.refreshInterval || 5
                })
              ]
            },
            {
              groupName: strings.CustomizationGroupName,
              groupFields: [
                PropertyPaneTextField('overdueColor', {
                  label: strings.OverdueColorLabel,
                  description: 'Hex color code (e.g., #d13438)',
                  placeholder: '#d13438'
                }),
                PropertyPaneTextField('inProgressColor', {
                  label: strings.InProgressColorLabel,
                  description: 'Hex color code (e.g., #0078d4)',
                  placeholder: '#0078d4'
                }),
                PropertyPaneTextField('completedColor', {
                  label: strings.CompletedColorLabel,
                  description: 'Hex color code (e.g., #107c10)',
                  placeholder: '#107c10'
                }),
                PropertyPaneToggle('showProgressBars', {
                  label: strings.ShowProgressBarsLabel,
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneToggle('showSourceBadges', {
                  label: strings.ShowSourceBadgesLabel,
                  onText: 'Show',
                  offText: 'Hide'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
