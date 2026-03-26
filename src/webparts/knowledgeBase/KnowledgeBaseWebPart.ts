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

import * as strings from 'KnowledgeBaseWebPartStrings';
import KnowledgeBase from './components/KnowledgeBase';
import { IKnowledgeBaseProps } from './components/IKnowledgeBaseProps';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface IKnowledgeBaseWebPartProps {
  title: string;
  listName: string;
  defaultCategory: string;
  itemsPerPage: number;
  enableSearch: boolean;
  enableCategoryFilter: boolean;
}

export default class KnowledgeBaseWebPart extends BaseClientSideWebPart<IKnowledgeBaseWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IKnowledgeBaseProps> = React.createElement(
      KnowledgeBase,
      {
        title: this.properties.title || 'OP Knowledge Base',
        listName: this.properties.listName || 'KnowledgeBase',
        defaultCategory: this.properties.defaultCategory || 'All',
        itemsPerPage: this.properties.itemsPerPage || 10,
        enableSearch: this.properties.enableSearch !== false,
        enableCategoryFilter: this.properties.enableCategoryFilter !== false,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    configureFluentUi();
    await super.onInit();
    return Promise.resolve();
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
              groupName: strings.GeneralGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                })
              ]
            },
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  placeholder: 'KnowledgeBase'
                })
              ]
            },
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneDropdown('defaultCategory', {
                  label: strings.DefaultCategoryFieldLabel,
                  options: [
                    { key: 'All', text: 'All' },
                    { key: 'HR', text: 'HR' },
                    { key: 'IT', text: 'IT' },
                    { key: 'Policies', text: 'Policies' },
                    { key: 'Finance', text: 'Finance' },
                    { key: 'General', text: 'General' }
                  ]
                }),
                PropertyPaneSlider('itemsPerPage', {
                  label: strings.ItemsPerPageFieldLabel,
                  min: 5,
                  max: 50,
                  step: 5,
                  showValue: true
                }),
                PropertyPaneToggle('enableSearch', {
                  label: strings.EnableSearchFieldLabel,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneToggle('enableCategoryFilter', {
                  label: strings.EnableCategoryFilterFieldLabel,
                  onText: 'Enabled',
                  offText: 'Disabled'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
