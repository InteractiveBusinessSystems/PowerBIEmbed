import { useCallback, useReducer, useState } from "react";
import * as config from "../config/authConfig";
import {AuthenticationContext} from "adal-node";
import { SPFI, spfi } from "@pnp/sp";
import { getSP } from "../config/PNPjsPresets";

export interface getAccessTokenInitialState {
  accessToken: string;
  getAccessTokenIsLoading: boolean;
  getAccessTokenError: unknown;
}

type Action = { type: "FETCH_START" } | { type: "FETCH_SUCCESS"; payload: getAccessTokenInitialState["accessToken"] } | { type: "FETCH_ERROR"; payload: getAccessTokenInitialState["getAccessTokenError"] } | { type: "RESET_getAccessToken" };

export const initialState: getAccessTokenInitialState = {
  accessToken: null,
  getAccessTokenIsLoading: false,
  getAccessTokenError: null,
};

const getAccessTokenReducer = (state: getAccessTokenInitialState, action: Action) => {
  switch (action.type) {
    case 'FETCH_START': {
      return { accessToken: null, getAccessTokenIsLoading: true, getAccessTokenError: null };
    }
    case 'FETCH_SUCCESS': {
      return { accessToken: action.payload, getAccessTokenIsLoading: false, getAccessTokenError: null };
    }
    case 'FETCH_ERROR': {
      return { accessToken: null, getAccessTokenIsLoading: false, getAccessTokenError: action.payload };
    }
    case 'RESET_getAccessToken': {
      return {
        accessToken: null,
        getAccessTokenIsLoading: false,
        getAccessTokenError: null
      };
    }
    default:
      return state;
  }
};

export const useGetAccessToken = () => {
  const [accessTokenState, getAccessTokenDispatch] = useReducer(getAccessTokenReducer, initialState);

  const getAccessToken = useCallback(async ()=> {
    getAccessTokenDispatch({type: "FETCH_START"});

    const context = new AuthenticationContext(config.authorityUrl);
    console.log(context);

    const tokenCallback = ((err, tokenResponse) => {
      if (err) {
        console.log(err);
        getAccessTokenDispatch({type: 'FETCH_ERROR', payload: err});
      }

      if(tokenResponse) {
        console.log(tokenResponse);
        getAccessTokenDispatch({type: 'FETCH_SUCCESS', payload: tokenResponse.accessToken});
      }
    });

    debugger;
    context.acquireTokenWithUsernamePassword(config.resource, config.MUUserName, config.MUpassword, config.clientId, tokenCallback);

  },[]);
  return { accessTokenState, getAccessToken };
};
