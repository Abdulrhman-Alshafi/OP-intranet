import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HeroSectionV2WebPartStrings';
import HeroSectionV2 from './components/HeroSectionV2';
import { IHeroSectionV2Props } from './components/IHeroSectionV2Props';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface IHeroSectionV2WebPartProps {
  heading: string;
  subheading: string;
  backgroundImage: string;
  ctaText: string;
  ctaLink: string;
  showCta: boolean;
  badgeText: string;
  overlayColor: string;
  overlayOpacity: number;
  heroHeight: number;
  textColor: string;
  textAlignment: 'left' | 'center' | 'right';
}

export default class HeroSectionV2WebPart extends BaseClientSideWebPart<IHeroSectionV2WebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IHeroSectionV2Props> = React.createElement(
      HeroSectionV2,
      {
        heading: this.properties.heading || '',
        subheading: this.properties.subheading || '',
        backgroundImage: this.properties.backgroundImage || '',
        ctaText: this.properties.ctaText || '',
        ctaLink: this.properties.ctaLink || '',
        showCta: this.properties.showCta !== false,
        badgeText: this.properties.badgeText || '',
        overlayColor: this.properties.overlayColor || '#000000',
        overlayOpacity: this.properties.overlayOpacity ?? 50,
        heroHeight: this.properties.heroHeight || 500,
        textColor: this.properties.textColor || '#ffffff',
        textAlignment: this.properties.textAlignment || 'center',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    configureFluentUi();
    await super.onInit();
    this._environmentMessage = await this._getEnvironmentMessage();
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
          header: { description: 'Configure the Hero Section V2' },
          groups: [
            {
              groupName: 'Content',
              groupFields: [
                PropertyPaneTextField('heading', {
                  label: 'Heading',
                  multiline: false
                }),
                PropertyPaneTextField('subheading', {
                  label: 'Subheading',
                  multiline: true,
                  rows: 3
                }),
                PropertyPaneTextField('badgeText', {
                  label: 'Badge Text (optional)',
                  description: 'Small label shown above the heading, e.g. "Welcome" or "News"',
                  multiline: false
                }),
                PropertyPaneTextField('textColor', {
                  label: 'Text Color (hex)',
                  description: 'Default is #ffffff (white)',
                  multiline: false
                }),
                PropertyPaneChoiceGroup('textAlignment', {
                  label: 'Text Alignment',
                  options: [
                    { key: 'left', text: 'Left' },
                    { key: 'center', text: 'Center' },
                    { key: 'right', text: 'Right' }
                  ]
                })
              ]
            },
            {
              groupName: 'Call to Action',
              groupFields: [
                PropertyPaneToggle('showCta', {
                  label: 'Show CTA Button',
                  onText: 'Visible',
                  offText: 'Hidden'
                }),
                PropertyPaneTextField('ctaText', {
                  label: 'Button Text',
                  multiline: false,
                  disabled: !this.properties.showCta
                }),
                PropertyPaneTextField('ctaLink', {
                  label: 'Button Link (URL)',
                  multiline: false,
                  disabled: !this.properties.showCta
                })
              ]
            },
            {
              groupName: 'Background',
              groupFields: [
                PropertyPaneTextField('backgroundImage', {
                  label: 'Background Image URL',
                  description: 'Paste a URL for the hero background image',
                  multiline: false
                }),
                PropertyPaneTextField('overlayColor', {
                  label: 'Overlay Color (hex)',
                  description: 'e.g. #000000 for black, #1a1a2e for dark blue',
                  multiline: false
                }),
                PropertyPaneSlider('overlayOpacity', {
                  label: 'Overlay Opacity (%)',
                  min: 0,
                  max: 100,
                  step: 5,
                  showValue: true
                }),
                PropertyPaneSlider('heroHeight', {
                  label: 'Hero Height (px)',
                  min: 300,
                  max: 800,
                  step: 10,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
