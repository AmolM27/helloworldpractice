import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldPracticeWebPart.module.scss';
import * as strings from 'HelloWorldPracticeWebPartStrings';

import {
	Environment, EnvironmentType
} from '@microsoft/sp-core-library';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IHelloWorldPracticeWebPartProps {
  description: string;
  color: string;
}

export default class HelloWorldPracticeWebPart extends BaseClientSideWebPart<IHelloWorldPracticeWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorldPractice }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
              <p class="ms-font-l ms-fontColor-white">Selected Color: ${escape(this.properties.color)}</p>
            </div>
          </div>
        </div>
      </div>
      <div id="lists">
      </div>
      `;

      this.getListsInfo();
  }

  protected get disableReactivePropertyChanges(): boolean
  {
    return true;
  }

  private getListsInfo()
  {
    let html: string = '';
    if (Environment.type === EnvironmentType.Local)
    {
          this.domElement.querySelector('#lists').innerHTML = "Sorry this does not work in local workbench";
    }
    else {
    this.context.spHttpClient.get(
    this.context.pageContext.web.absoluteUrl + '/_api/web/lists?$filter=Hidden eq false',
    SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      response.json().then((listObjects: any) => {
        listObjects.value.forEach(listObject => {
          html += `
          <ul>
            <li>
              <span class="ms-font-l">${listObject.Title}</span>
            </li>
          </ul>
          `
        });
        this.domElement.querySelector('#lists').innerHTML = html;
      })
    });
  }
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
                }),
                PropertyPaneDropdown('color', {
                  label: 'Dropdown',
                    options: [
                      { key: '1', text: 'Red'},
                      { key: '2', text: 'Blue'},
                      { key: '3', text: 'Green'},
                    ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
