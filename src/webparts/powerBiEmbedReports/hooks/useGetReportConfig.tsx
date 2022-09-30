import { useCallback, useReducer } from "react";
import { AadHttpClient, AadHttpClientFactory, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import * as msal from "@azure/msal-browser";
import { getSP } from '../config/PNPjsPresets';
import { spfi, SPFI } from '@pnp/sp';
import { IReportsList } from "./IReportsList.types";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { forEach } from "lodash";

export interface IReportConfig {
  EmbedToken: string,
  EmbedUrl: string,
  AccessToken: string,
  ReportId: string,
}

export interface getReportConfigInitialState {
  ReportConfig: IReportsList[];
  getReportConfigIsLoading: boolean;
  getReportConfigError: unknown;
}

type Action = { type: "FETCH_START" } | { type: "FETCH_SUCCESS"; payload: getReportConfigInitialState["ReportConfig"] } | { type: "FETCH_ERROR"; payload: getReportConfigInitialState["getReportConfigError"] } | { type: "RESET_getReportConfig" };

export const initialState: getReportConfigInitialState = {
  ReportConfig: [{ ReportName: "", DataSetsId: "", WorkspaceId: "", ReportId: "", ReportUrl: "", ViewerType: "", UsersWhoCanView: [], Id: undefined, EmbedToken: "", EmbedUrl: "", AccessToken: "" }],
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
        ReportConfig: [{ ReportName: "", WorkspaceId: "", ReportId: "", ReportUrl: "", ViewerType: "", UsersWhoCanView: [], Id: undefined, EmbedToken: "", EmbedUrl: "", AccessToken: "" }],
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

  const msalConfig = {
    auth: {
      clientId: '170af556-d26c-40b3-9a96-361ce11d683d',
      authority: 'https://login.microsoftonline.com/4ec55493-6b1c-4565-a868-2ae940882c82',
      // redirectUri: 'http://localhost:3000/blank.html'
      redirectUri: 'http://localhost:3000/myapp'
    }
  };




  const getReportConfig = useCallback(async (aadHttpClient: AadHttpClientFactory, reports) => {
    getReportConfigDispatch({ type: "FETCH_START" });
    let results: IReportsList[];
    // let requestOptions: IHttpClientOptions;

    // // const AzureFunctionUrl = 'https://powerbisharepoint.azurewebsites.net/api/GetToken?code=epmSB0fCWciQLsZ9R2AiaDBRw-eJYaA5wVH5npnw28liAzFudxQ9KQ==';
    // const AzureFunctionUrl = 'https://maryvill-test-function.azurewebsites.net/api/gettoken?code=oZN_P3a2R2zRuTBegu-rla57DETCkRJQFOlva3M_pv0EAzFua53Pew%3D%3D';

    // console.log(reports);
    // let reportsString = JSON.stringify(reports);
    // console.log(reportsString);

    // const requestUrl = AzureFunctionUrl;
    // console.log(requestUrl);
    // const httpClientOptions: IHttpClientOptions = {
    //   body: reportsString
    // }

    // aadHttpClient.getClient(
    //   //This is the App's Client ID
    //   // 'a0ac7405-5171-4909-8f8f-c195ab60c28d'
    //   '170af556-d26c-40b3-9a96-361ce11d683d'
    // )
    //   .then((client: AadHttpClient): void => {
    //     client.post(
    //       //Calling AzureFunction
    //       requestUrl,
    //       AadHttpClient.configurations.v1,
    //       httpClientOptions
    //     )
    //       .then((response: HttpClientResponse): Promise<IReportsList[]> => {
    //         if (response.status === 200) {
    //           return response.json();
    //         }
    //         else {
    //           console.log(response);
    //           throw "Token fetch request failed!";
    //         }
    //       })
    //       .then((jsonResponse: IReportsList[]): void => {
    //         console.log(jsonResponse);
    //         getReportConfigDispatch({ type: 'FETCH_SUCCESS', payload: jsonResponse });
    //       })
    //       .catch(error => {
    //         console.log(error);
    //         getReportConfigDispatch({ type: 'FETCH_ERROR', payload: error });
    //       });
    //   })
    //   .catch(aadError => {
    //     console.log(aadError);
    //     getReportConfigDispatch({ type: 'FETCH_ERROR', payload: aadError });
    //   });

    console.log(reports);
    let loginResponse;
    const currentUser:any = await spfi(sp).web.currentUser();
    const currentUserEmail:string = currentUser.Email;
    console.log(currentUserEmail);

    const silentRequest = {
      scopes: ["https://analysis.windows.net/powerbi/api/Report.Read.All"],
      loginHint: currentUserEmail
    };

    let accessToken = "";


    const msalInstance = new msal.PublicClientApplication(msalConfig);

    try {
      loginResponse = await msalInstance.ssoSilent(silentRequest);
      console.log(loginResponse);
    }
    catch (err) {
      if (err instanceof InteractionRequiredAuthError) {
        msalInstance.loginPopup(silentRequest).then(response =>{
          console.log(response);
          loginResponse = response;
        }).catch(error => {
          console.log(error);
          getReportConfigDispatch({type: 'FETCH_ERROR', payload: error});
        });
      }
      else {
        console.log(err);
        getReportConfigDispatch({type: 'FETCH_ERROR', payload: err});
      }
    }

    // try {
    //   const loginResponse = await msalInstance.loginPopup(silentRequest);
    //   console.log(loginResponse);
    // }
    // catch (err) {
    //   console.log(err);
    //   getReportConfigDispatch({type: 'FETCH_ERROR', payload: err});
    // }


    if(loginResponse) {
        // get access token silently
        msalInstance.acquireTokenSilent(silentRequest).then((response) => {
            //get access token from response
            accessToken = response.accessToken;

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

            //Power BI REST API calls to get the embed URLs of the reports
            reports.forEach(report => {
                    fetch("https://api.powerbi.com/v1.0/myorg/groups/" + report.WorkspaceId + "/reports/" + report.ReportId, {
                        headers: {
                          "Authorization": "Bearer " + accessToken
                        },
                        method: "Get"
                    }).then(response => {
                        const errorMessage: string[] = [];
                        errorMessage.push("Error occurred while fetching the embed URL of the report")
                        errorMessage.push("Request Id: " + response.headers.get("requestId"));

                        response.json().then(body => {
                          if(response.ok){
                            console.log(body["embedUrl"]);
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
                          }
                          else {
                            errorMessage.push("Error " + response.status + ": " + body.error.code);
                            console.log(errorMessage);
                            getReportConfigDispatch({type: 'FETCH_ERROR', payload: errorMessage});
                          }
                        })
                        .catch(jsonError => {
                            errorMessage.push("Error " + response.status + ": An error has occurred");
                            console.log(errorMessage);
                            getReportConfigDispatch({type: 'FETCH_ERROR', payload: errorMessage});
                        });
                    })
                    .catch(embedError => {
                        console.log(embedError);
                        getReportConfigDispatch({type: 'FETCH_ERROR', payload: embedError});
                    });

                    return results;
            });
            getReportConfigDispatch({type: 'FETCH_SUCCESS', payload: results});
        })
        .catch((mError: msal.AuthError) => {
          if(mError.name === "InteractionRequiredAuthError"){
            msalInstance.acquireTokenPopup(silentRequest);
            // May need to add the Power BI Calls (173-247) here too
          }
          else {
            getReportConfigDispatch({type: 'FETCH_ERROR', payload: mError});
          }
        });
    }
    else {
      // user is not logged in or cached, you will need to log them in to acquire a token
      msalInstance.loginPopup(silentRequest);
    }

  }, []);
  return { reportConfigState, getReportConfig };
};

