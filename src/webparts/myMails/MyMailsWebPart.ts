import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MyMailsWebPartStrings';
import MyMails from './components/MyMails';
import { IMyMailsProps } from './components/IMyMailsProps';

export interface IMyMailsWebPartProps {
  applicationID: string;
  redirectUri: string;
  tenantUrl: string;
}

export default class MyMailsWebPart extends BaseClientSideWebPart <IMyMailsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMyMailsProps> = React.createElement(
      MyMails,
      {        
        applicationID: this.properties.applicationID,
        redirectUri: this.properties.redirectUri,
        tenantUrl: this.properties.tenantUrl,
        httpClient: this.context.httpClient,
        userMail: this.context.pageContext.user.email
      }
    );

    ReactDom.render(element, this.domElement);
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('applicationID', {
                  label: strings.ApplicationIDFieldLabel
                }),
                PropertyPaneTextField('redirectUri', {
                  label: strings.RedirectUriFieldLabel
                }),
                PropertyPaneTextField('tenantUrl', {
                  label: strings.TenantUrlFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
