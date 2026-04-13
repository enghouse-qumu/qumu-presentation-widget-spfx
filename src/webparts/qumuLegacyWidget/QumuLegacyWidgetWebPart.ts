import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './QumuLegacyWidgetWebPart.module.scss';
import * as strings from 'QumuLegacyWidgetWebPartStrings';

interface IKvWidgetGlobal {
  widget: (config: {
    guid: string;
    type: 'featured' | 'vertical' | 'carousel' | 'playback' | 'thumbnail' | 'grid' | 'playlist',
    selector: string;
  }) => void;
}

export interface IQumuLegacyWidgetWebPartProps {
  guid: string;
  host: string;
  type: 'featured' | 'vertical' | 'carousel' | 'playback' | 'thumbnail' | 'grid' | 'playlist',
}

export default class QumuLegacyWidgetWebPart extends BaseClientSideWebPart<IQumuLegacyWidgetWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `<div id="${this._getWidgetElementId()}" class="${styles.widgetHost}"></div>`;

    this._renderWidget().catch(() => undefined);
  }

  protected onDispose(): void {
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
                PropertyPaneDropdown('type', {
                  label: strings.TypeFieldLabel,
                  options: [
                    {
                      key: 'featured',
                      text: 'featured',
                    },
                    {
                      key: 'vertical',
                      text: 'vertical',
                    },
                    {
                      key: 'carousel',
                      text: 'carousel',
                    },
                    {
                      key: 'playback',
                      text: 'playback',
                    },
                    {
                      key: 'thumbnail',
                      text: 'thumbnail',
                    },
                    {
                      key: 'grid',
                      text: 'grid',
                    },
                    {
                      key: 'playlist',
                      text: 'playlist',
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
    try {
      const kv = await SPComponentLoader.loadScript<IKvWidgetGlobal>(
        new URL('widgets/application-spfx.js', this.properties.host).toString(),
        {
          globalExportsName: 'KV'
        }
      );

      kv.widget({
        guid: this.properties.guid,
        type: this.properties.type,
        selector: `#${this._getWidgetElementId()}`
      });
    } catch (error) {
      console.error(strings.LoadFailedPrefix, error instanceof Error ? error.message : strings.UnknownError);
    }
  }

  private _getWidgetElementId(): string {
    return `legacy-widget-${this.instanceId}`;
  }
}
