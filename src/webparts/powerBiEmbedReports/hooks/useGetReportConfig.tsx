import { useCallback, useReducer } from "react";
import { AadHttpClient, AadHttpClientFactory, HttpClientResponse } from '@microsoft/sp-http';
import * as config from "../config/authConfig";
import { IReportsList } from "./IReportsList.types";

// export interface IReportsConfig {
//   ReportName: string;
//   WorkspaceId: string;
//   ReportId: string;
//   ReportSectionId: string;
//   ReportUrl: string;
//   ViewerType: string;
//   UsersWhoCanView: [];
//   Id: number;
//   EmbedToken: string,
//   EmbedUrl: string,
//   AccessToken: string,
// }

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
  ReportConfig: [{ReportName: "", WorkspaceId: "", ReportId: "", ReportSectionId: "", ReportUrl: "", ViewerType: "", UsersWhoCanView: [], Id: undefined, EmbedToken: "", EmbedUrl: "", AccessToken: ""}],
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
        ReportConfig: [{ReportName: "", WorkspaceId: "", ReportId: "", ReportSectionId: "", ReportUrl: "", ViewerType: "", UsersWhoCanView: [], Id: undefined, EmbedToken: "", EmbedUrl: "", AccessToken: ""}] ,
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

  const getReportConfig = useCallback(async (aadHttpClient: AadHttpClientFactory, reports)=> {
    getReportConfigDispatch({type: "FETCH_START"});
    let results: IReportsList[];

    reports.forEach(report => {
      aadHttpClient.getClient(
        //This is the App's Client ID
        '170af556-d26c-40b3-9a96-361ce11d683d'
      )
      .then((client:AadHttpClient): void =>{
        console.log(client);
        client.get(
          //This is the Azure Function Url
          'https://maryvillepowerbipocfunctionapp.azurewebsites.net/api/GetEmbedToken?code=XH2KY-GZomKT_jPvotQ6ADtLC2nyNFEqQSzJPZCjg7eiAzFuHNhhFw==',
          AadHttpClient.configurations.v1
        ).then((response: HttpClientResponse): Promise<IReportConfig> => {
          console.log(response);
          return response.json();
        })
        .then((jsonResponse: IReportConfig): void => {
          console.log(jsonResponse);
          results.push({
            "ReportName": report.ReportName,
            "WorkspaceId": report.WorkspaceId,
            "ReportId": report.ReportId,
            "ReportSectionId": report.ReportSectionId,
            "ReportUrl": report.ReportUrl,
            "ViewerType": report.ViewerType,
            "UsersWhoCanView": report.UsersWhoCanView,
            "Id": report.Id,
            "EmbedToken": jsonResponse.EmbedToken,
            "EmbedUrl": jsonResponse.EmbedUrl,
            "AccessToken": jsonResponse.AccessToken
          });

        })
        .catch(error => {
          console.log(error);
          getReportConfigDispatch({type: 'FETCH_ERROR', payload: error});
        });
      });

      getReportConfigDispatch({type: 'FETCH_SUCCESS', payload: results});

    })



  },[]);
  return { reportConfigState, getReportConfig };
};