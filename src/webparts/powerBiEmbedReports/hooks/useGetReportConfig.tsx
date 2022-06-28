import { useCallback, useReducer } from "react";
import { AadHttpClient, AadHttpClientFactory, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { IReportsList } from "./IReportsList.types";

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

  const getReportConfig = useCallback(async (aadHttpClient: AadHttpClientFactory, reports) => {
    getReportConfigDispatch({ type: "FETCH_START" });
    let results: IReportsList[];
    let requestOptions: IHttpClientOptions;
    const AzureFunctionUrl = 'https://powerbisharepoint.azurewebsites.net/api/GetToken?code=NT8rk9oNqDIVe92KsxQPodRrynS_4t9FeTxQK7WeNtPKAzFu-vzy8w==';

    let reportsString = JSON.stringify(reports);

    const requestUrl = AzureFunctionUrl;
    const httpClientOptions: IHttpClientOptions = {
      body: reportsString
    }

    aadHttpClient.getClient(
      //This is the App's Client ID
      'a0ac7405-5171-4909-8f8f-c195ab60c28d'
    )
      .then((client: AadHttpClient): void => {
        client.post(
          //Calling AzureFunction
          requestUrl,
          AadHttpClient.configurations.v1,
          httpClientOptions
        )
          .then((response: HttpClientResponse): Promise<IReportsList[]> => {
            if (response.status === 200) {
              return response.json();
            }
            else {
              throw "Token fetch request failed!";
            }
          })
          .then((jsonResponse: IReportsList[]): void => {
            getReportConfigDispatch({ type: 'FETCH_SUCCESS', payload: jsonResponse });
          })
          .catch(error => {
            console.log(error);
            getReportConfigDispatch({ type: 'FETCH_ERROR', payload: error });
          });
      })
      .catch(aadError => {
        console.log(aadError);
        getReportConfigDispatch({ type: 'FETCH_ERROR', payload: aadError });
      });

  }, []);
  return { reportConfigState, getReportConfig };
};
