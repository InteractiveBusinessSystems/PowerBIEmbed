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
    // const AzureFunctionUrl = 'https://powerbisharepoint.azurewebsites.net/api/GetToken?code=epmSB0fCWciQLsZ9R2AiaDBRw-eJYaA5wVH5npnw28liAzFudxQ9KQ==';
    const AzureFunctionUrl = 'https://maryvill-test-function.azurewebsites.net/api/gettoken?code=oZN_P3a2R2zRuTBegu-rla57DETCkRJQFOlva3M_pv0EAzFua53Pew%3D%3D';

    console.log(reports);
    let reportsString = JSON.stringify(reports);
    console.log(reportsString);

    const requestUrl = AzureFunctionUrl;
    console.log(requestUrl);
    const httpClientOptions: IHttpClientOptions = {
      body: reportsString
    }

    aadHttpClient.getClient(
      //This is the App's Client ID
      // 'a0ac7405-5171-4909-8f8f-c195ab60c28d'
      '170af556-d26c-40b3-9a96-361ce11d683d'
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
              console.log(response);
              throw "Token fetch request failed!";
            }
          })
          .then((jsonResponse: IReportsList[]): void => {
            console.log(jsonResponse);
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
