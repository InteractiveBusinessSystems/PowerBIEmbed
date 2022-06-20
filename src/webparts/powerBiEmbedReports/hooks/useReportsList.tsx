import { useReducer, useCallback} from 'react';
import { IReportsList } from './IReportsList.types';
import {getSP, getGraph} from '../config/PNPjsPresets';
import { spfi, SPFI } from '@pnp/sp';
import { graphfi, GraphFI } from '@pnp/graph';

export interface reportsListInitialState {
  reports: IReportsList[];
  reportsListIsLoading: boolean;
  reportsListError: unknown;
}

type Action = {type: "FETCH_START"} | {type: "FETCH_SUCCESS"; payload: reportsListInitialState["reports"]} | {type: "FETCH_ERROR"; payload: reportsListInitialState["reportsListError"]} | {type: "RESET_REPORTSLIST"};

export const initialState: reportsListInitialState = {
  reports: null,
  reportsListIsLoading: false,
  reportsListError: null,
};

const reportsListReducer = (state: reportsListInitialState, action: Action) => {
  switch(action.type) {
    case 'FETCH_START': {
      return { reports: null, reportsListIsLoading: true, reportsListError: null };
    }
    case 'FETCH_SUCCESS': {
      return { reports: action.payload, reportsListIsLoading: false, reportsListError: null };
    }
    case 'FETCH_ERROR': {
      return { reports: null, reportsListIsLoading: false, reportsListError: action.payload };
    }
    case 'RESET_REPORTSLIST': {
      return { reports: null,
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
  const graph: GraphFI = getGraph();

  const getReportsListResults = useCallback(async () => {
      reportsListDispatch({type: "FETCH_START"});
      let results:IReportsList[] = [];

      const currentUser:any = await spfi(sp).web.currentUser();
      const currentUserGroups = await graphfi(graph).me.getMemberGroups(true);

      try{
      const items: any[] = await spfi(sp).web.lists.getByTitle('Power BI Reports List').items.select('Title', 'Id', 'DataSetsId', 'WorkspaceId', 'ReportId', 'ReportSectionId', 'ReportUrl', 'ViewerType', 'UsersWhoCanView/Name').expand('UsersWhoCanView').top(500)();

        items.forEach((report) => {
          if(report.ViewerType === 'Group'){
            let contains = false;
            let usersWhoCanView = report.UsersWhoCanView;

            usersWhoCanView.forEach((group)=> {
              const gName = group.Name;
              const groupName = gName.substring(14);
              let contains = false;

              currentUserGroups.forEach((userGroup) => {
                if (groupName.toLowerCase() === userGroup.toLowerCase()) {
                  contains = true;
                }
              });
            });

            if (contains) {
              results.push({
                "ReportName": report.Title,
                "DataSetsId": report.DataSetsId,
                "WorkspaceId": report.WorkspaceId,
                "ReportId": report.ReportId,
                "ReportSectionId": report.ReportSectionId,
                "ReportUrl": report.ReportUrl,
                "ViewerType": report.ViewerType,
                "UsersWhoCanView": report.UsersWhoCanView,
                "Id": parseInt(report.Id)
              });
            }
          }

          if(report.ViewerType === 'User'){
            let contains = false;
            let usersWhoCanView = report.UsersWhoCanView;

            usersWhoCanView.forEach((user)=> {
              let userName = user.Name;
              let userEmail = userName.substring(18);

              if(userEmail.toLowerCase() === currentUser.Email.toLowerCase()){
                contains = true;
              }
            });

            if (contains) {
              results.push({
                "ReportName": report.Title,
                "DataSetsId": report.DataSetsId,
                "WorkspaceId": report.WorkspaceId,
                "ReportId": report.ReportId,
                "ReportSectionId": report.ReportSectionId,
                "ReportUrl": report.ReportUrl,
                "ViewerType": report.ViewerType,
                "UsersWhoCanView": report.UsersWhoCanView,
                "Id": parseInt(report.Id)
              });
            }

          }

        });
        reportsListDispatch({type: 'FETCH_SUCCESS', payload: results});
      }
      catch(e){
        console.log(e.message);
        reportsListDispatch({type: 'FETCH_ERROR', payload: e.message});
      }

  },[]);
  return {state, getReportsListResults, reportsListDispatch};
};
