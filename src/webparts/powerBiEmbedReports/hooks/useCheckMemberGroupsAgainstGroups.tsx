import { useCallback, useEffect, useReducer, useState } from 'react';

export interface checkMemberGroupsInitialState {
  data: boolean;
  checkMemberGroupsIsLoading: boolean;
  checkMemberGroupsError: unknown;
}

type Action = { type: "FETCH_START" } | { type: "FETCH_SUCCESS"; payload: checkMemberGroupsInitialState["data"] } | { type: "FETCH_ERROR"; payload: checkMemberGroupsInitialState["checkMemberGroupsError"] } | { type: "RESET_checkMemberGroups" };

export const initialState: checkMemberGroupsInitialState = {
  data: false,
  checkMemberGroupsIsLoading: false,
  checkMemberGroupsError: null,
};

const checkMemberGroupsReducer = (state: checkMemberGroupsInitialState, action: Action) => {
  switch (action.type) {
    case 'FETCH_START': {
      return { data: null, checkMemberGroupsIsLoading: true, checkMemberGroupsError: null };
    }
    case 'FETCH_SUCCESS': {
      return { data: action.payload, checkMemberGroupsIsLoading: false, checkMemberGroupsError: null };
    }
    case 'FETCH_ERROR': {
      return { data: null, checkMemberGroupsIsLoading: false, checkMemberGroupsError: action.payload };
    }
    case 'RESET_checkMemberGroups': {
      return {
        data: false,
        checkMemberGroupsIsLoading: false,
        checkMemberGroupsError: null
      };
    }
    default:
      return state;
  }
};

export const useCheckMemberGroupsAgainstGroups = () => {
  const [state, checkMemberGroupsDispatch] = useReducer(checkMemberGroupsReducer, initialState);

  const checkMemberGroupsAgainstGroups = useCallback((groups, userGroups) => {
    checkMemberGroupsDispatch({ type: "FETCH_START" });
    const audienceGroups = groups;
    const currentUserGroups = userGroups;

    if (audienceGroups && currentUserGroups) {
      audienceGroups.forEach(audienceGroup => {
        let contains = false;
        const group = audienceGroup.id;
        const audienceGroupName = group.substring(14);
        currentUserGroups.forEach(userGroup => {
          if (audienceGroupName === userGroup) {
            contains = true;
          }
        });
        if (contains) {
          checkMemberGroupsDispatch({ type: 'FETCH_SUCCESS', payload: true });
        }
        else {
          checkMemberGroupsDispatch({ type: 'FETCH_SUCCESS', payload: false });
        }
      });
    }
    else{
      checkMemberGroupsDispatch({type: 'FETCH_ERROR', payload: 'Please select a security group to audience your web part'});
    }

  }, []);
  return { state, checkMemberGroupsAgainstGroups };
}

