import {useState, useReducer, useCallback} from 'react';
import { IReportsList } from './IReportsList.types';
import {getSP} from '../config/PNPjsPresets';
// import "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";

export interface reportsListInitialState {
  data: IReportsList[];
  reportsListIsLoading: boolean;
  reportsListError: unknown;
}

type Action = {type: "FETCH_START"} | {type: "FETCH_SUCCESS"; payload: reportsListInitialState["data"]} | {type: "FETCH_ERROR"; payload: reportsListInitialState["reportsListError"]} | {type: "RESET_REPORTSLIST"};

export const initialState: reportsListInitialState = {
  data: [{ReportName: "", WorkspaceId: "", ReportId: "", ReportSectionId: "", Department: "", Id: undefined}],
  reportsListIsLoading: false,
  reportsListError: null,
};

const reportsListReducer = (state: reportsListInitialState, action: Action) => {
  switch(action.type) {
    case 'FETCH_START': {
      return { data: null, reportsListIsLoading: true, reportsListError: null };
    }
    case 'FETCH_SUCCESS': {
      return { data: action.payload, reportsListIsLoading: false, reportsListError: null };
    }
    case 'FETCH_ERROR': {
      return { data: null, reportsListIsLoading: false, reportsListError: action.payload };
    }
    case 'RESET_REPORTSLIST': {
      return { data: [{ReportName: "", WorkspaceId: "", ReportId: "", ReportSectionId: "", Department: "", Id: undefined}],
      reportsListIsLoading: false,
      reportsListError: null};
    }
    default:
      return state;
  }
};

export const useReportsList = () => {
  const[state, reportsListDispatch] = useReducer(reportsListReducer, initialState);
  const sp = getSP();

  const getReportsListResults = useCallback(async () => {
      reportsListDispatch({type: "FETCH_START"});
      let results:IReportsList[] = [];
      const items: any[] = await sp.web.lists.getByTitle('Power BI Reports List').items.select('Title', 'Id', 'WorkspaceId', 'ReportId', 'ReportSectionId', 'Department').top(500)();
      items.forEach((report) => {
        results.push({
          "ReportName": report.Title,
          "WorkspaceId": report.WorkspaceId,
          "ReportId": report.ReportId,
          "ReportSectionId": report.ReportSectionId,
          "Department": report.Department,
          "Id": parseInt(report.Id)
        });

      });

      return results;

  },[]);
  return {state, getReportsListResults, reportsListDispatch};
};
