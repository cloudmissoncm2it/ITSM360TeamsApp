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
import {Itsm360TeamsConfigure} from './components/Itsm360TeamsConfigure';
import { IItsm360TeamsConfigureProps } from './components/Itsm360TeamsConfigure';
import * as microsoftTeams from '@microsoft/teams-js';

export interface IItsm360TeamsAppWebPartProps {
  description: string;
}

export default class Itsm360TeamsAppWebPart extends BaseClientSideWebPart<IItsm360TeamsAppWebPartProps> {
  
  private _teamsContext:microsoftTeams.Context;
 
  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public render(): void {
    const element: React.ReactElement<IItsm360TeamsConfigureProps > = React.createElement(
      Itsm360TeamsConfigure,
      {
        context:this.context,
        teamscontext:this._teamsContext
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
