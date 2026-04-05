import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'SalesPerformanceDashboardWebPartStrings';
import SalesPerformanceDashboard from './components/SalesPerformanceDashboard';
import { ISalesPerformanceDashboardProps } from './components/ISalesPerformanceDashboardProps';

export interface ISalesPerformanceDashboardWebPartProps {
  listId: string;
  chartType: string;
}

export default class SalesPerformanceDashboardWebPart extends BaseClientSideWebPart<ISalesPerformanceDashboardWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;

  protected onInit(): Promise<void> {
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeVariant = this._themeProvider.tryGetTheme();
    if (this._themeProvider && this._themeProvider.themeChangedEvent) {
      this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
    }
    this._isDarkTheme = !!this._themeVariant && !!this._themeVariant.isInverted;
    return super.onInit();
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this._isDarkTheme = !!args.theme && !!args.theme.isInverted;
    this.render();
  }

  public render(): void {
    const element: React.ReactElement<ISalesPerformanceDashboardProps> = React.createElement(
      SalesPerformanceDashboard,
      {
        listId: this.properties.listId,
        chartType: this.properties.chartType,
        isDarkTheme: this._isDarkTheme,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyFieldListPicker('listId', {
                  label: 'Select a data source',
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  onGetErrorMessage: undefined,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                }),
                PropertyPaneDropdown('chartType', {
                  label: strings.ChartTypeFieldLabel,
                  options: [
                    { key: 'bar', text: 'Bar Chart' },
                    { key: 'line', text: 'Line Chart' }
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
