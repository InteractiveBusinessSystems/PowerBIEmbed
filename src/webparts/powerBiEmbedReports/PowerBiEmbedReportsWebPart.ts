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
import * as config from "./config/authConfig";
import { UserAgentApplication, AuthError, AuthResponse } from "msal";
import { SPFI, spfi } from '@pnp/sp';
import { ISiteUser, ISiteUserInfo } from '@pnp/sp/site-users/types';


export interface IPowerBiEmbedReportsWebPartProps {
  description?: string;
  groups?: IPropertyFieldGroupOrPerson[];
  userGroups?: string[];
  accessToken?: string;
  accessTokenError?: string;
}

export default class PowerBiEmbedReportsWebPart extends BaseClientSideWebPart<IPowerBiEmbedReportsWebPartProps> {
  currentUser: ISiteUserInfo;
  userGroups: string[];
  accessToken: string;
  accessTokenError: string;

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
    if (!this.renderedOnce) {

      const ssoRequest = {
        loginHint: this.currentUser.Email
      };

      const msalInstance: UserAgentApplication = new UserAgentApplication(config.msalConfig);

      console.log(`msalInstance: ${msalInstance}`);

      // Power BI REST API call to refresh User Permissions in Power BI
      // Refreshes user permissions and makes sure the user permissions are fully updated
      // https://docs.microsoft.com/rest/api/power-bi/users/refreshuserpermissions
      const tryRefreshUserPermissions = (): void => {
        fetch("https://api.powerbi.com/v1.0/myorg/RefreshUserPermissions", {
          headers: {
            "Authorization": "Bearer " + this.accessToken
          },
          method: "POST"
        })
          .then(response => {
            if (response.ok) {
              console.log("User permissions refreshed successfully.");
            } else {
              // Too many requests in one hour will cause the API to fail
              if (response.status === 429) {
                console.error("Permissions refresh will be available in up to an hour.");
              } else {
                console.error(response);
              }
            }
          })
          .catch(refreshError => {
            console.error("Failure in making API call." + refreshError);
          });
      };

      const successCallback = (response: AuthResponse): void => {
        // if (response.tokenType === "id_token") {
        //   useGetAccessToken();
        // } else
        if (response.tokenType === "access_token") {
          console.log(`successCallbackresponse: ${response}`);
          this.accessToken = response.accessToken;
          tryRefreshUserPermissions();
        } else {
          this.accessTokenError = `Token type is: ${response.tokenType}`;
        }
      };

      const failCallback = (failError: AuthError): void => {
        this.accessTokenError = `Redirect error: ${failError}`;
      };

      msalInstance.handleRedirectCallback(successCallback, failCallback);

      //check if there is a cached user
      if (msalInstance.getAccount()) {
        // get access token silently from cached id-token

        msalInstance.acquireTokenSilent(config.loginRequest)
        // msalInstance.ssoSilent(ssoRequest)
          .then((response: AuthResponse) => {
            console.log(`aquireTokenSilentResponse: ${response}`);
            //get access token from response: response.accessToken
            this.accessToken = response.accessToken;
          })
          .catch((err: AuthError) => {
            //refresh access token silently from cached id-token
            //makes the call to handleredirectcallback
            if (err.name === "InteractionRequiredAuthError") {
              msalInstance.acquireTokenRedirect(config.loginRequest);
            }
            else {
              this.accessTokenError = err.toString();
            }
          });
      } else {
        //user is not logged in or cached, we need to log them in to acquire a token
        msalInstance.loginRedirect(config.loginRequest);
      }
    }

    const element: React.ReactElement<IPowerBiEmbedReportsWebPartProps> = React.createElement(
      PowerBiEmbedReports,
      {
        // description: this.properties.description
        groups: this.properties.groups,
        userGroups: this.userGroups,
        accessToken: this.accessToken,
        accessTokenError: this.accessTokenError,
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
