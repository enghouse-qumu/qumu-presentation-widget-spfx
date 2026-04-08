import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  type WidgetOptions,
  PresentationWidget,
} from '@qumu/widgets';
import '@qumu/widgets/presentation-widget.css';

import styles from './QumuPresentationWidgetWebPart.module.scss';
import * as strings from 'QumuPresentationWidgetWebPartStrings';

export interface IQumuPresentationWidgetWebPartProps {
  guid: string;
  host: string;
  playbackMode: WidgetOptions['playbackMode'];
}

export default class QumuPresentationWidgetWebPart extends BaseClientSideWebPart<IQumuPresentationWidgetWebPartProps> {
  private _widget?: PresentationWidget;

  public render(): void {
    this.domElement.innerHTML = `<div id="${this._getWidgetElementId()}" class="${styles.widgetHost}"></div>`;

    this._renderWidget().catch(() => undefined);
  }

  protected onDispose(): void {
    this._widget?.destroy();
    this._widget = undefined;
    this.domElement.innerHTML = '';
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
                PropertyPaneTextField('host', {
                  label: strings.HostFieldLabel
                }),
                PropertyPaneTextField('guid', {
                  label: strings.GuidFieldLabel
                }),
                PropertyPaneDropdown('playbackMode', {
                  label: strings.PlaybackModeFieldLabel,
                  options: [
                    {
                      key: 'inline',
                      text: 'inline',
                    },
                    {
                      key: 'inline-autoload',
                      text: 'inline-autoload',
                    },
                    {
                      key: 'inline-autoplay',
                      text: 'inline-autoplay',
                    },
                    {
                      key: 'modal',
                      text: 'modal',
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private async _renderWidget(): Promise<void> {
    this._widget?.destroy();
    this._widget = undefined;

    try {
      this._widget = await PresentationWidget.create({
        host: this.properties.host,
        guid: this.properties.guid,
        selector: `#${this._getWidgetElementId()}`,
        widgetOptions: {
          playbackMode: this.properties.playbackMode
        }
      });
    } catch (error) {
      console.error(strings.LoadFailedPrefix, error instanceof Error ? error.message : strings.UnknownError);
    }
  }

  private _getWidgetElementId(): string {
    return `widget-${this.instanceId}`;
  }
}
