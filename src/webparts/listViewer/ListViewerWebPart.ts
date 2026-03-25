import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import ListViewer from './components/ListViewer';
import { IListViewerProps } from './components/IListViewerProps';

export interface IListViewerWebPartProps {
  listId: string;
  maxItems: number;
  viewStyle: 'grid' | 'list' | 'table';
  title: string;
}

export default class ListViewerWebPart extends BaseClientSideWebPart<IListViewerWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IListViewerProps> = React.createElement(
      ListViewer,
      {
        listId: this.properties.listId,
        maxItems: this.properties.maxItems || 10,
        viewStyle: this.properties.viewStyle || 'grid',
        title: this.properties.title || 'List Viewer',
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        themeVariant: this._isDarkTheme ? 'dark' : 'light'
      }
    );

    ReactDom.render(element, this.domElement);
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
          header: { description: 'Configure ListView Premium.' },
          groups: [
            {
              groupName: 'Data Source',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title'
                }),
                PropertyFieldListPicker('listId', {
                  label: 'Select a list',
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
                })
              ]
            },
            {
              groupName: 'Configuration',
              groupFields: [
                PropertyPaneSlider('maxItems', {
                  label: 'Maximum items to show',
                  min: 1,
                  max: 50,
                  step: 1
                }),
                PropertyPaneDropdown('viewStyle', {
                  label: 'Display Style',
                  options: [
                    { key: 'grid', text: 'Card Grid Layout' },
                    { key: 'list', text: 'Compact List Layout' },
                    { key: 'table', text: 'Data Table Layout' }
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
