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
import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { graphfi, GraphFI } from '@pnp/graph';

export interface IPowerBiEmbedReportsWebPartProps {
  description?: string;
  groups?: IPropertyFieldGroupOrPerson[];
}

export default class PowerBiEmbedReportsWebPart extends BaseClientSideWebPart<IPowerBiEmbedReportsWebPartProps> {
  isAudienced: boolean = false;


  protected async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
    getGraph(this.context);

    const graph: GraphFI = getGraph();
    const currentUserGroups:any = await graphfi(graph).me.getMemberGroups(true);
    const audienceGroups: any = this.properties.groups;

    if(audienceGroups){
      audienceGroups.forEach(audienceGroup => {
        let contains = false;
        const group = audienceGroup.id;
        const audienceGroupName = group.substring(14);
        currentUserGroups.forEach(userGroup => {
            if(audienceGroupName === userGroup){
              contains = true;
            }
        });
        if(contains){
          this.isAudienced = true;
        }
        else {
          this.isAudienced = false;
        }
      });
    }

  }

  public render(): void {

   const element: React.ReactElement<IPowerBiEmbedReportsWebPartProps> = React.createElement(
      PowerBiEmbedReports,
      {
        // description: this.properties.description
        groupIds: this.properties.groups,
        isAudienced: this.isAudienced
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
