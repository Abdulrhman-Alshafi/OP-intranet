import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme, ThemeChangedEventArgs, ThemeProvider } from '@microsoft/sp-component-base';

import { HelpDeskDashboard } from './components/HelpDeskDashboard';
import { IHelpDeskDevicesProps } from './components/IHelpDeskDevicesProps';
import { configureFluentUi } from '../../common/configureFluentUi';
import { HelpDeskDeviceService } from '../../services/HelpDeskDeviceService';

export interface IHelpDeskDevicesWebPartProps {
  devicesListName: string;
  requestsListName: string;
  helpDeskGroupName: string;
}

export default class HelpDeskDevicesWebPart extends BaseClientSideWebPart<IHelpDeskDevicesWebPartProps> {
  private _themeProvider: ThemeProvider | undefined;
  private _themeVariant: IReadonlyTheme | undefined;

  private readonly _handleThemeChanged = (themeEventArgs: ThemeChangedEventArgs): void => {
    this._themeVariant = themeEventArgs.theme;
    this.render();
  };

  public render(): void {
    const service = new HelpDeskDeviceService(
      this.context.pageContext.web.absoluteUrl,
      this.context.spHttpClient
    );

    const element: React.ReactElement<IHelpDeskDevicesProps> = React.createElement(
      HelpDeskDashboard,
      {
        service,
        devicesListName: this.properties.devicesListName || 'HelpDeskDevices',
        requestsListName: this.properties.requestsListName || 'HelpDeskDeviceRequests',
        helpDeskGroupName: this.properties.helpDeskGroupName || 'Help Desk',
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
          header: { description: 'Help Desk Configuration' },
          groups: [
            {
              groupName: 'General Settings',
              groupFields: [
                PropertyPaneTextField('devicesListName', {
                  label: 'Devices List Name',
                  description: 'Internal name of the Devices SharePoint list.'
                }),
                PropertyPaneTextField('requestsListName', {
                  label: 'Requests List Name',
                  description: 'Internal name of the Requests SharePoint list.'
                }),
                PropertyPaneTextField('helpDeskGroupName', {
                  label: 'Help Desk Group Name',
                  description: 'Exact name of the SharePoint group granted access.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
