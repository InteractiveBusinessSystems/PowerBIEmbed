import { useReducer, useCallback} from 'react';
import { IReportsList } from './IReportsList.types';
import {getSP, getGraph} from '../config/PNPjsPresets';
import { spfi, SPFI } from '@pnp/sp';
import { graphfi, GraphFI } from '@pnp/graph';
import { useCheckUserGroup } from './useCheckUserGroup';

export interface reportsListInitialState {
  data: IReportsList[];
  reportsListIsLoading: boolean;
  reportsListError: unknown;
}

type Action = {type: "FETCH_START"} | {type: "FETCH_SUCCESS"; payload: reportsListInitialState["data"]} | {type: "FETCH_ERROR"; payload: reportsListInitialState["reportsListError"]} | {type: "RESET_REPORTSLIST"};

export const initialState: reportsListInitialState = {
  data: [{ReportName: "", WorkspaceId: "", ReportId: "", ReportSectionId: "", ReportUrl: "", ViewerType: "", UsersWhoCanView: [], Id: undefined}],
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
      return { data: [{ReportName: "", WorkspaceId: "", ReportId: "", ReportSectionId: "", Department: "", UsersWhoCanView: [], ViewerType: "", Id: undefined}],
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
      const currentUserGroups:any = await graphfi(graph).me.getMemberGroups(true);

      try{
      const items: any[] = await spfi(sp).web.lists.getByTitle('Power BI Reports List').items.select('Title', 'Id', 'WorkspaceId', 'ReportId', 'ReportSectionId', 'ReportUrl', 'ViewerType', 'UsersWhoCanView/Name').expand('UsersWhoCanView').top(500)();

        items.forEach((report) => {
          if(report.ViewerType === 'Group'){
            let contains = false;
            let usersWhoCanView = report.UsersWhoCanView;

            usersWhoCanView.forEach( (group)=> {
              contains = useCheckUserGroup(group, currentUserGroups);
            });

            if (contains) {
              results.push({
                "ReportName": report.Title,
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
