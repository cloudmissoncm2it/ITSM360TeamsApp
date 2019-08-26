import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'Itsm360TeamsAppWebPartStrings';
import {Itsm360TeamsApp} from './components/Itsm360TeamsApp';
import { IItsm360TeamsAppProps } from './components/Itsm360TeamsApp';

export interface IItsm360TeamsAppWebPartProps {
  description: string;
}

export default class Itsm360TeamsAppWebPart extends BaseClientSideWebPart<IItsm360TeamsAppWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IItsm360TeamsAppProps > = React.createElement(
      Itsm360TeamsApp,
      {
        description: this.properties.description,
        sphttpclient:this.context.spHttpClient,
        currentuser:this.context.pageContext.user,
        context:this.context
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
