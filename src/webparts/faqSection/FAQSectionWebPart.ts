import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';

import * as strings from 'FAQSectionWebPartStrings';
import { FAQSection } from './components/FAQSection';
import { IFAQSectionProps } from './components/IFAQSectionProps';
import { configureFluentUi } from '../../common/configureFluentUi';
import { FAQService } from '../../services/FAQService';

export interface IFAQSectionWebPartProps {
  listName: string;
}

export default class FAQSectionWebPart extends BaseClientSideWebPart<IFAQSectionWebPartProps> {
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;

  private readonly _handleThemeChanged = (themeEventArgs: ThemeChangedEventArgs): void => {
    this._themeVariant = themeEventArgs.theme;
    this.render();
  };

  public render(): void {
    const service = new FAQService(
      this.context.pageContext.web.absoluteUrl,
      this.context.spHttpClient,
      this.properties.listName || 'FAQs'
    );

    const element: React.ReactElement<IFAQSectionProps> = React.createElement(FAQSection, {
      service,
      themeVariant: this._themeVariant
    });

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    configureFluentUi();
    await super.onInit();

    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);
    this._themeVariant = this._themeProvider.tryGetTheme();
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChanged);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    if (this._themeProvider) {
      this._themeProvider.themeChangedEvent.remove(this, this._handleThemeChanged);
    }
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
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  description: 'Internal name of the SharePoint list storing FAQs.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
