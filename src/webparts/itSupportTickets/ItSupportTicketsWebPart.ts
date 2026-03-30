import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IReadonlyTheme,
  ThemeChangedEventArgs,
  ThemeProvider
} from '@microsoft/sp-component-base';
import { HttpClient } from '@microsoft/sp-http';

import * as strings from 'ItSupportTicketsWebPartStrings';
import { TicketDashboard } from './components/TicketDashboard';
import { IItSupportTicketsProps } from './components/IItSupportTicketsProps';
import { ITopTicketService } from '../../services/ITopTicketService';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface IItSupportTicketsWebPartProps {
  functionBaseUrl: string;
  functionKey: string;
  itAdminGroupName: string;
}

export default class ItSupportTicketsWebPart extends BaseClientSideWebPart<IItSupportTicketsWebPartProps> {
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;

  private readonly _handleThemeChanged = (args: ThemeChangedEventArgs): void => {
    this._themeVariant = args.theme;
    this.render();
  };

  public render(): void {
    const httpClient: HttpClient = this.context.httpClient;

    const service = new ITopTicketService(
      this.properties.functionBaseUrl || '',
      this.properties.functionKey || '',
      httpClient
    );

    const element: React.ReactElement<IItSupportTicketsProps> = React.createElement(
      TicketDashboard,
      {
        service,
        isConfigured: !!(this.properties.functionBaseUrl),
        itAdminGroupName: this.properties.itAdminGroupName || 'IT Support',
        currentUserEmail: this.context.pageContext.user.email,
        currentUserLoginName: this.context.pageContext.user.loginName,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
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
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.GeneralGroupName,
              groupFields: [
                PropertyPaneTextField('functionBaseUrl', {
                  label: strings.FunctionBaseUrlFieldLabel,
                  description: strings.FunctionBaseUrlFieldDesc,
                  placeholder: 'https://my-func-app.azurewebsites.net'
                }),
                PropertyPaneTextField('functionKey', {
                  label: strings.FunctionKeyFieldLabel,
                  description: strings.FunctionKeyFieldDesc
                }),
                PropertyPaneTextField('itAdminGroupName', {
                  label: strings.ItAdminGroupFieldLabel,
                  description: strings.ItAdminGroupFieldDesc
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
