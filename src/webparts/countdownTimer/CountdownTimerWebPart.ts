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

import { PropertyFieldDateTimePicker, DateConvention, TimeConvention, IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';

import * as strings from 'CountdownTimerWebPartStrings';
import CountdownTimer from './components/CountdownTimer';
import { ICountdownTimerProps } from './components/ICountdownTimerProps';
import { configureFluentUi } from '../../common/configureFluentUi';

export interface ICountdownTimerWebPartProps {
  eventTitle: string;
  eventDescription: string;
  targetDate: IDateTimeFieldValue;
  backgroundImage: string;
  overlayOpacity: number;
  layout: 'horizontal' | 'compact' | 'banner';
  showDays: boolean;
  showHours: boolean;
  showMinutes: boolean;
  showSeconds: boolean;
  completedMessage: string;
  accentColor: string;
}

export default class CountdownTimerWebPart extends BaseClientSideWebPart<ICountdownTimerWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<ICountdownTimerProps> = React.createElement(
      CountdownTimer,
      {
        eventTitle: this.properties.eventTitle || '',
        eventDescription: this.properties.eventDescription || '',
        targetDate: this.properties.targetDate && this.properties.targetDate.value
          ? new Date(this.properties.targetDate.value).toISOString()
          : '',
        backgroundImage: this.properties.backgroundImage || '',
        overlayOpacity: this.properties.overlayOpacity !== undefined ? this.properties.overlayOpacity : 50,
        layout: this.properties.layout || 'horizontal',
        showDays: this.properties.showDays !== false,
        showHours: this.properties.showHours !== false,
        showMinutes: this.properties.showMinutes !== false,
        showSeconds: this.properties.showSeconds !== false,
        completedMessage: this.properties.completedMessage || 'The event has started!',
        accentColor: this.properties.accentColor || '#0078d4',
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    configureFluentUi();
    await super.onInit();
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Configure your Countdown Timer' },
          groups: [
            {
              groupName: 'Event Details',
              groupFields: [
                PropertyPaneTextField('eventTitle', { label: 'Event Title' }),
                PropertyPaneTextField('eventDescription', { label: 'Event Description', multiline: true }),
                PropertyFieldDateTimePicker('targetDate', {
                  label: 'Target Date & Time',
                  initialDate: this.properties.targetDate,
                  dateConvention: DateConvention.DateTime,
                  timeConvention: TimeConvention.Hours24,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'targetDateField'
                }),
                PropertyPaneTextField('completedMessage', { label: 'Completed Message' })
              ]
            },
            {
              groupName: 'Appearance',
              groupFields: [
                PropertyPaneDropdown('layout', {
                  label: 'Layout',
                  options: [
                    { key: 'horizontal', text: 'Horizontal (Default)' },
                    { key: 'compact', text: 'Compact' },
                    { key: 'banner', text: 'Banner (Side by Side)' }
                  ]
                }),
                PropertyPaneTextField('accentColor', {
                  label: 'Accent Color (Hex)',
                  description: 'e.g. #0078d4, #107c10, #d83b01'
                }),
                {
                  targetProperty: 'backgroundImage',
                  type: 1,
                  properties: {
                    key: 'backgroundImageField',
                    onRender: (elem: HTMLElement) => {
                      // Don't recreate if already rendered
                      if (elem.querySelector('[data-bg-field]')) {
                        // Update preview if value changed externally
                        const preview = elem.querySelector('img') as HTMLImageElement;
                        if (preview && this.properties.backgroundImage) {
                          preview.src = this.properties.backgroundImage;
                          preview.style.display = 'block';
                        }
                        return;
                      }

                      const wrapper = document.createElement('div');
                      wrapper.setAttribute('data-bg-field', 'true');
                      wrapper.style.marginBottom = '8px';

                      const label = document.createElement('label');
                      label.textContent = 'Background Image';
                      label.style.cssText = 'display:block;font-weight:600;font-size:14px;margin-bottom:5px;color:#323130;font-family:"Segoe UI",sans-serif';
                      wrapper.appendChild(label);

                      // URL text input
                      const urlInput = document.createElement('input');
                      urlInput.type = 'text';
                      urlInput.placeholder = 'Paste image URL...';
                      urlInput.value = this.properties.backgroundImage && !this.properties.backgroundImage.startsWith('data:image') ? this.properties.backgroundImage : '';
                      urlInput.style.cssText = 'width:100%;padding:6px 8px;border:1px solid #8a8886;border-radius:2px;font-size:14px;font-family:"Segoe UI",sans-serif;box-sizing:border-box;margin-bottom:6px';
                      urlInput.addEventListener('change', (e: Event) => {
                        const oldVal = this.properties.backgroundImage;
                        const newVal = (e.target as HTMLInputElement).value;
                        this.properties.backgroundImage = newVal;
                        this.onPropertyPaneFieldChanged('backgroundImage', oldVal, newVal);
                        this.render();
                        // Update preview
                        const p = wrapper.querySelector('img') as HTMLImageElement;
                        if (p) { p.src = newVal; p.style.display = newVal ? 'block' : 'none'; }
                      });
                      wrapper.appendChild(urlInput);

                      // "OR" label
                      const orLabel = document.createElement('span');
                      orLabel.textContent = '— OR Upload File —';
                      orLabel.style.cssText = 'display:block;font-size:12px;color:#605e5c;margin-bottom:6px';
                      wrapper.appendChild(orLabel);

                      // File upload
                      const fileInput = document.createElement('input');
                      fileInput.type = 'file';
                      fileInput.accept = 'image/*';
                      fileInput.style.cssText = 'margin-bottom:8px';
                      fileInput.addEventListener('change', (e: Event) => {
                        const file = (e.target as HTMLInputElement).files![0];
                        if (file) {
                          const reader = new FileReader();
                          reader.onload = (event: any) => {
                            const oldVal = this.properties.backgroundImage;
                            const newVal = event.target.result;
                            this.properties.backgroundImage = newVal;
                            this.onPropertyPaneFieldChanged('backgroundImage', oldVal, newVal);
                            this.render();
                            const p = wrapper.querySelector('img') as HTMLImageElement;
                            if (p) { p.src = newVal; p.style.display = 'block'; }
                          };
                          reader.readAsDataURL(file);
                        }
                      });
                      wrapper.appendChild(fileInput);

                      // Preview
                      const preview = document.createElement('img');
                      preview.style.cssText = 'height:60px;width:auto;max-width:100%;object-fit:cover;border-radius:2px;border:1px solid #edebe9';
                      preview.style.display = this.properties.backgroundImage ? 'block' : 'none';
                      if (this.properties.backgroundImage) preview.src = this.properties.backgroundImage;
                      wrapper.appendChild(preview);

                      elem.appendChild(wrapper);
                    },
                    onDispose: () => { /* noop */ }
                  }
                },
                PropertyPaneSlider('overlayOpacity', {
                  label: 'Background Overlay Darkness',
                  min: 0,
                  max: 100,
                  step: 5,
                  showValue: true
                })
              ]
            },
            {
              groupName: 'Display Units',
              groupFields: [
                PropertyPaneToggle('showDays', { label: 'Show Days' }),
                PropertyPaneToggle('showHours', { label: 'Show Hours' }),
                PropertyPaneToggle('showMinutes', { label: 'Show Minutes' }),
                PropertyPaneToggle('showSeconds', { label: 'Show Seconds' })
              ]
            }
          ]
        }
      ]
    };
  }
}
