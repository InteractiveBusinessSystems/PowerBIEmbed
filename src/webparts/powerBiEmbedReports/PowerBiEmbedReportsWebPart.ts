import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PowerBiEmbedReportsWebPartStrings';
import PowerBiEmbedReports from './components/PowerBiEmbedReports';
import { IPowerBiEmbedReportsProps } from './components/IPowerBiEmbedReportsProps';

import { getSP, getGraph } from './config/PNPjsPresets';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

export interface IPowerBiEmbedReportsWebPartProps {
  description?: string;
  groups?: IPropertyFieldGroupOrPerson[];
}

export default class PowerBiEmbedReportsWebPart extends BaseClientSideWebPart <IPowerBiEmbedReportsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPowerBiEmbedReportsProps> = React.createElement(
      PowerBiEmbedReports,
      {
        // description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected async onInit(): Promise<void> {

    await super.onInit();
    getSP(this.context);
    getGraph(this.context);
  }

  constructor() {
    super();
    Object.defineProperty(this, "dataVersion", {
      get() {
        return Version.parse('1.0');
      }
    });
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Use this web part to audience and embed PowerBI reports'
          },
          groups: [
            {
              // groupName: strings.BasicGroupName,
              groupFields: [
                // PropertyPaneTextField('description', {
                //   label: strings.DescriptionFieldLabel
                // }),
                // PropertyFieldPeoplePicker('groups', {
                //   label: 'Target Audience',
                //   initialData: this.properties.groups,
                //   allowDuplicate: false,
                //   principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                //   onPropertyChange: this.onPropertyPaneFieldChanged,
                //   context: this.context,
                //   properties: this.properties,
                //   onGetErrorMessage: null,
                //   deferredValidationTime: 0,
                //   key: 'groupFieldId'
                // })
              ]
            }
          ]
        }
      ]
    };
  }
}
