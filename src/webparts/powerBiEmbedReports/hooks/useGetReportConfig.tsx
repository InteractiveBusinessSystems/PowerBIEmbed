import { useCallback, useReducer } from "react";
import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import * as config from "../config/authConfig";

export interface IReportConfig {
  EmbedToken: string,
  EmbedUrl: string,
  ReportId: string,
  AccessToken: string,
}

export interface getReportConfigInitialState {
  ReportConfig: IReportConfig;
  getReportConfigIsLoading: boolean;
  getReportConfigError: unknown;
}

type Action = { type: "FETCH_START" } | { type: "FETCH_SUCCESS"; payload: getReportConfigInitialState["ReportConfig"] } | { type: "FETCH_ERROR"; payload: getReportConfigInitialState["getReportConfigError"] } | { type: "RESET_getReportConfig" };

export const initialState: getReportConfigInitialState = {
  ReportConfig: {
    EmbedToken: null,
    EmbedUrl: null,
    ReportId: null,
    AccessToken: null
  },
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

  const getReportConfig = useCallback(async (aadHttpClient: AadHttpClientFactory)=> {
    getReportConfigDispatch({type: "FETCH_START"});
    let results: IReportConfig;

    aadHttpClient.getClient(
      //The ID at the end of this Url is the App's Client ID
      'https://ibsmtg.onmicrosoft.com/170af556-d26c-40b3-9a96-361ce11d683d'
    )
    .then((client:AadHttpClient) =>{
      console.log(client);
      client.get(
        //This is the Azure Function Url
        'https://maryvillepowerbipocfunctionapp.azurewebsites.net/api/GetEmbedToken?code=XH2KY-GZomKT_jPvotQ6ADtLC2nyNFEqQSzJPZCjg7eiAzFuHNhhFw==',
        AadHttpClient.configurations.v1
      ).then(response => {
        console.log(response);
        return response.json();
      })
      .then(jsonResponse => {
        console.log(jsonResponse);
        results ={
          "EmbedToken": jsonResponse.EmbedToken,
          "EmbedUrl": jsonResponse.EmbedUrl,
          "ReportId": jsonResponse.ReportId,
          "AccessToken": jsonResponse.AccessToken
        };

        getReportConfigDispatch({type: 'FETCH_SUCCESS', payload: results});
      })
      .catch(error => {
        console.log(error);
        getReportConfigDispatch({type: 'FETCH_ERROR', payload: error});
      });
    });

  },[]);
  return { reportConfigState, getReportConfig };
};
