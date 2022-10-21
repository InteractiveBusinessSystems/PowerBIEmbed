import { useCallback, useReducer } from "react";
import { AadHttpClient, AadHttpClientFactory, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import * as msal from "@azure/msal-browser";
import { getSP } from '../config/PNPjsPresets';
import { spfi, SPFI } from '@pnp/sp';
import { IReportsList } from "./IReportsList.types";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { forEach } from "lodash";
import * as React from "react";
import { PowerBIEmbed } from "powerbi-client-react";
import { models } from 'powerbi-client';
import styles from '../components/PowerBiEmbedReports.module.scss';

export interface getReportConfigInitialState {
  // ReportConfig: IReportsList[];
  ReportConfig: any[];
  getReportConfigIsLoading: boolean;
  getReportConfigError: unknown;
}

type Action = { type: "FETCH_START" } | { type: "FETCH_SUCCESS"; payload: getReportConfigInitialState["ReportConfig"] } | { type: "FETCH_ERROR"; payload: getReportConfigInitialState["getReportConfigError"] } | { type: "RESET_getReportConfig" };

export const initialState: getReportConfigInitialState = {
  ReportConfig: null,
  getReportConfigIsLoading: false,
  getReportConfigError: null,
};

const getReportConfigReducer = (state: getReportConfigInitialState, action: Action) => {
  switch (action.type) {
    case 'FETCH_START': {
      return { ReportConfig: null, getReportConfigIsLoading: true, getReportConfigError: null };
    }
    case 'FETCH_SUCCESS': {
      return { ReportConfig: action.payload, getReportConfigIsLoading: false, getReportConfigError: null };
    }
    case 'FETCH_ERROR': {
      return { ReportConfig: null, getReportConfigIsLoading: false, getReportConfigError: action.payload };
    }
    case 'RESET_getReportConfig': {
      return {
        ReportConfig: null,
        getReportConfigIsLoading: false,
        getReportConfigError: null
      };
    }
    default:
      return state;
  }
};

export const useGetReportConfig = () => {
  const [reportConfigState, getReportConfigDispatch] = useReducer(getReportConfigReducer, initialState);
  const sp: SPFI = getSP();

  // const msalConfig = {
  //   auth: {
  //     clientId: '170af556-d26c-40b3-9a96-361ce11d683d',
  //     authority: 'https://login.microsoftonline.com/4ec55493-6b1c-4565-a868-2ae940882c82',
  //   }
  // };

  //Maryville
  const msalConfig = {
    auth: {
      clientId: '26fd5c26-2f1b-449f-9cce-0b27ecc9016e',
      authority: 'https://login.microsoftonline.com/0a89b54f-c81a-40aa-91d6-91a54133911e',
    }
  };

  const getReportConfig = useCallback(async (aadHttpClient: AadHttpClientFactory, reports) => {
    getReportConfigDispatch({ type: "FETCH_START" });
    let reportsLength = reports.length;
    let results: IReportsList[] = [];
    let loginResponse;
    let accessToken = "";
    const currentUser: any = await spfi(sp).web.currentUser();
    const currentUserEmail: string = currentUser.Email;

    const silentRequest = {
      scopes: ["https://analysis.windows.net/powerbi/api/Report.Read.All"],
      loginHint: currentUserEmail
    };

    // Create MsalInstance
    const msalInstance = new msal.PublicClientApplication(msalConfig);

    // Check if users is signedIn using SSO
    try {
      loginResponse = await msalInstance.ssoSilent(silentRequest);
    }
    catch (err) {
      if (err instanceof InteractionRequiredAuthError) {
        msalInstance.loginPopup(silentRequest).then(response => {
          console.log(response);
          loginResponse = response;
        }).catch(error => {
          console.log(error);
          getReportConfigDispatch({ type: 'FETCH_ERROR', payload: error });
        });
      }
      else {
        console.log(err);
        getReportConfigDispatch({ type: 'FETCH_ERROR', payload: err });
      }
    }

    if (loginResponse) {

      //get access token from response
      accessToken = loginResponse.accessToken;

      // Power BI REST API call to refresh User Permissions in Power BI
      // Refreshes user permissions and makes sure the user permissions are fully updated
      // https://docs.microsoft.com/rest/api/power-bi/users/refreshuserpermissions
      fetch("https://api.powerbi.com/v1.0/myorg/RefreshUserPermissions", {
        headers: {
          "Authorization": "Bearer " + accessToken
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
        .catch(error => {
          console.error("Failure in making API call." + error);
        });

      let reportsCounter = 0;
      //Power BI REST API calls to get the embed URLs of the reports
      reports.forEach(report => {

        fetch("https://api.powerbi.com/v1.0/myorg/groups/" + report.WorkspaceId + "/reports/" + report.ReportId, {
          headers: {
            "Authorization": "Bearer " + accessToken
          },
          method: "Get"
        }).then(response => {
          const errorMessage: string[] = [];

          response.json().then(body => {
            if (response.ok) {
              let embedUrl = body["embedUrl"];

              results.push({
                "ReportName": report.ReportName,
                "ReportId": report.ReportId,
                "Id": report.Id,
                "DataSetsId": report.DatasetsId,
                "WorkspaceId": report.WorkspaceId,
                "ReportUrl": report.ReportUrl,
                "UsersWhoCanView": report.UsersWhoCanView,
                "ViewerType": report.ViewerType,
                "AccessToken": accessToken,
                "EmbedUrl": embedUrl
              });

              reportsCounter ++;
              if(reportsCounter === reportsLength){
                console.log(results);
                getReportConfigDispatch({ type: 'FETCH_SUCCESS', payload: results });
              }

            }
            else {
              errorMessage.push("Error " + response.status + ": " + body.error.code);
              console.log(errorMessage);
              getReportConfigDispatch({ type: 'FETCH_ERROR', payload: errorMessage });
            }
          })
            .catch(jsonError => {
              errorMessage.push("Error " + response.status + ": An error has occurred");
              console.log(errorMessage);
              getReportConfigDispatch({ type: 'FETCH_ERROR', payload: errorMessage });
            });
        })
          .catch(embedError => {
            console.log(embedError);
            getReportConfigDispatch({ type: 'FETCH_ERROR', payload: embedError });
          });
      });
    }
    else {
      // user is not logged in or cached, you will need to log them in to acquire a token
      msalInstance.loginPopup(silentRequest);
    }

  }, []);
  return { reportConfigState, getReportConfig };
};

