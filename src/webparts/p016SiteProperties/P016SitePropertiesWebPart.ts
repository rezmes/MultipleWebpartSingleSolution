import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './P016SitePropertiesWebPart.module.scss';
import * as strings from 'P016SitePropertiesWebPartStrings';

export interface IP016SitePropertiesWebPartProps {
  description: string;
}

export default class P016SitePropertiesWebPart extends BaseClientSideWebPart<IP016SitePropertiesWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.p016SiteProperties }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>

              <p class="${ styles.description }">${escape(this.context.pageContext.web.absoluteUrl)}</p>
                            <p class="${ styles.description }">${escape(this.context.pageContext.web.title)}</p>
                                          <p class="${ styles.description }">${escape(this.context.pageContext.web.serverRelativeUrl)}</p>
                                                        <p class="${ styles.description }">${escape(this.context.pageContext.user.displayName)}</p>


              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}