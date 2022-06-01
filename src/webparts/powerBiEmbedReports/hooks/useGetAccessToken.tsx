import { useCallback, useReducer } from "react";
import * as config from "../config/authConfig";
import { UserAgentApplication, AuthError, AuthResponse } from "msal";

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

    const msalInstance: UserAgentApplication = new UserAgentApplication(config.msalConfig);
    console.log(msalInstance);

    const account = msalInstance.getAccount();
    console.log(account);

    const silentRequest = {
      scopes: config.scopes,
      account: account,
      forceRefresh: false
    };

    msalInstance.handleRedirectCallback((error: AuthError, response: AuthResponse) => {
      console.log('Redirect Callback was called');
      if(response){
        console.log(response);
        getAccessTokenDispatch({type: 'FETCH_SUCCESS', payload: response.accessToken});
      }
      else {
        console.log(error);
        getAccessTokenDispatch({type: 'FETCH_ERROR', payload: error});
      }
    });

    msalInstance.acquireTokenSilent(silentRequest)
    .then(response => {
      console.log(response);
      getAccessTokenDispatch({type: 'FETCH_SUCCESS', payload: response.accessToken});
    })
    .catch(error => {
      console.log(error);
      return msalInstance.acquireTokenRedirect(silentRequest);
      // getAccessTokenDispatch({type: 'FETCH_ERROR', payload: error});
    });

  },[]);
  return { accessTokenState, getAccessToken };

  // const msalInstance: UserAgentApplication = new UserAgentApplication(config.msalConfig);

  // console.log(`msalInstance: ${msalInstance}`);

  // // Power BI REST API call to refresh User Permissions in Power BI
  //   // Refreshes user permissions and makes sure the user permissions are fully updated
  //   // https://docs.microsoft.com/rest/api/power-bi/users/refreshuserpermissions
  //   const tryRefreshUserPermissions = (): void => {
  //     fetch("https://api.powerbi.com/v1.0/myorg/RefreshUserPermissions", {
  //         headers: {
  //             "Authorization": "Bearer " + accessToken
  //         },
  //         method: "POST"
  //     })
  //     .then(response => {
  //         if (response.ok) {
  //             console.log("User permissions refreshed successfully.");
  //         } else {
  //             // Too many requests in one hour will cause the API to fail
  //             if (response.status === 429) {
  //                 console.error("Permissions refresh will be available in up to an hour.");
  //             } else {
  //                 console.error(response);
  //             }
  //         }
  //     })
  //     .catch(refreshError => {
  //         console.error("Failure in making API call." + refreshError);
  //     });
  // };

  // const successCallback = (response: AuthResponse): void => {
  //     if(response.tokenType === "id_token") {
  //        getAccessToken();
  //     } else if (response.tokenType === "access_token") {
  //         console.log(`successCallbackresponse: ${response}`);
  //         setAccessToken(response.accessToken);
  //         tryRefreshUserPermissions();
  //     } else {
  //       setAccessTokenError(`Token type is: ${response.tokenType}`);
  //     }
  // };

  // const failCallback = (failError: AuthError): void => {
  //   setAccessTokenError(`Redirect error: ${failError}`);
  // };

  // msalInstance.handleRedirectCallback(successCallback,failCallback);

  // console.log("just before get token");
  // // //check if there is a cached user
  // // if (msalInstance.getAccount()) {
  //   //get access token silently from cached id-token
  //   msalInstance.acquireTokenSilent(config.loginRequest)
  //     .then((response:AuthResponse) => {
  //       console.log(`aquireTokenSilentResponse: ${response}`);
  //       //get access token from response: response.accessToken
  //       setAccessToken(response.accessToken);
  //     })
  //     .catch((err: AuthError) => {
  //       //refresh access token silently from cached id-token
  //       //makes the call to handleredirectcallback
  //       console.log(err);
  //       if(err.name === "InteractionRequiredAuthError") {
  //         msalInstance.acquireTokenRedirect(config.loginRequest);
  //       }
  //       else {
  //         setAccessTokenError(err.toString());
  //       }
  //     });
  // // } else {
  // //   //user is not logged in or cached, we need to log them in to acquire a token
  // //   msalInstance.loginRedirect(config.loginRequest);
  // // }
  // return {accessToken, accessTokenError};
};
