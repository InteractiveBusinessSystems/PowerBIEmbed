import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PowerBiEmbedReportsWebPartStrings';
import PowerBiEmbedReports from './components/PowerBiEmbedReports';
import { IPowerBiEmbedReportsProps } from './components/IPowerBiEmbedReportsProps';

import { getSP, getGraph } from './config/PNPjsPresets';
import { GraphFI, graphfi } from '@pnp/graph';
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { SPFI, spfi } from '@pnp/sp';
import { ISiteUserInfo } from '@pnp/sp/site-users/types';


export interface IPowerBiEmbedReportsWebPartProps {
  description?: string;
  groups?: IPropertyFieldGroupOrPerson[];
  userGroups?: string[];
}

export default class PowerBiEmbedReportsWebPart extends BaseClientSideWebPart<IPowerBiEmbedReportsWebPartProps> {
  currentUser: ISiteUserInfo;
  userGroups: string[];

  protected async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
    getGraph(this.context);

    const graph: GraphFI = getGraph();
    this.userGroups = await graphfi(graph).me.getMemberGroups(true);

    const sp: SPFI = getSP();
    this.currentUser = await spfi(sp).web.currentUser();
  }

  public render(): void {

     const element: React.ReactElement<IPowerBiEmbedReportsWebPartProps> = React.createElement(
      PowerBiEmbedReports,
      {
        // description: this.properties.description
        groups: this.properties.groups,
        userGroups: this.userGroups,
        aadHttpClient: this.context.aadHttpClientFactory
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                PropertyFieldPeoplePicker('groups', {
                  label: 'Target Audience',
                  initialData: this.properties.groups,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'groupFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
