import { useReducer, useCallback} from 'react';
import { IReportsList } from './IReportsList.types';
import {getSP} from '../config/PNPjsPresets';
import { spfi, SPFI } from '@pnp/sp';

export interface reportsListInitialState {
  data: IReportsList[];
  reportsListIsLoading: boolean;
  reportsListError: unknown;
}

type Action = {type: "FETCH_START"} | {type: "FETCH_SUCCESS"; payload: reportsListInitialState["data"]} | {type: "FETCH_ERROR"; payload: reportsListInitialState["reportsListError"]} | {type: "RESET_REPORTSLIST"};

export const initialState: reportsListInitialState = {
  data: [{ReportName: "", WorkspaceId: "", ReportId: "", ReportSectionId: "", UsersWhoCanView: [], Id: undefined}],
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
  const sp: SPFI = getSP();

  const getReportsListResults = useCallback(async () => {
      reportsListDispatch({type: "FETCH_START"});
      let results:IReportsList[] = [];

      const user:any = await spfi(sp).web.currentUser();
      console.log(user);

      try{
      const items: any[] = await spfi(sp).web.lists.getByTitle('Power BI Reports List').items.select('Title', 'Id', 'WorkspaceId', 'ReportId', 'ReportSectionId', 'ViewerType', 'UsersWhoCanView/Name', 'UsersWhoCanView/FirstName', 'UsersWhoCanView/LastName', 'UsersWhoCanView/JobTitle', 'UsersWhoCanView/Department').expand('UsersWhoCanView').top(500)();

        items.forEach((report) => {
          results.push({
            "ReportName": report.Title,
            "WorkspaceId": report.WorkspaceId,
            "ReportId": report.ReportId,
            "ReportSectionId": report.ReportSectionId,
            "UsersWhoCanView": report.UsersWhoCanView,
            "Id": parseInt(report.Id)
          });

        });
        console.log(results);
        reportsListDispatch({type: 'FETCH_SUCCESS', payload: results});
      }
      catch(e){
        console.log(e.message);
        reportsListDispatch({type: 'FETCH_ERROR', payload: e.message});
      }

  },[]);
  return {state, getReportsListResults, reportsListDispatch};
};
