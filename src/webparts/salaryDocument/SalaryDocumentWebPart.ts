import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';

import * as strings from 'SalaryDocumentWebPartStrings';
import { SalaryDocumentWebPartDashboard } from './components/Dashboard';
import { ISalaryDocumentProps } from './components/ISalaryDocumentProps';
import { configureFluentUi } from '../../common/configureFluentUi';
import { SalaryDocumentService } from '../../services/SalaryDocumentService';

export interface ISalaryDocumentWebPartProps {
  libraryName: string;
  accountantGroupName: string;
}

export default class SalaryDocumentWebPart extends BaseClientSideWebPart<ISalaryDocumentWebPartProps> {
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;

  private readonly _handleThemeChanged = (themeEventArgs: ThemeChangedEventArgs): void => {
    this._themeVariant = themeEventArgs.theme;
    this.render();
  };

  public render(): void {
    const service = new SalaryDocumentService(
      this.context.pageContext.web.absoluteUrl,
      this.context.spHttpClient
    );

    const element: React.ReactElement<ISalaryDocumentProps> = React.createElement(
      SalaryDocumentWebPartDashboard,
      {
        service,
        libraryName: this.properties.libraryName || 'SalaryDocuments',
        accountantGroupName: this.properties.accountantGroupName || 'Accountants',
        currentUserEmail: this.context.pageContext.user.email,
        currentUserLoginName: this.context.pageContext.user.loginName,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        msGraphClientFactory: this.context.msGraphClientFactory,
        themeVariant: this._themeVariant
      }
    );

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
                PropertyPaneTextField('libraryName', {
                  label: strings.SalaryLibraryFieldLabel,
                  description: 'Internal name of the SharePoint document library.'
                }),
                PropertyPaneTextField('accountantGroupName', {
                  label: strings.AccountantGroupFieldLabel,
                  description: 'Exact name of the SharePoint group granted accountant access.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
