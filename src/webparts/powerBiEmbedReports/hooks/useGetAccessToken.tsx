import * as React from "react";
import { useState } from "react";
import * as config from "../config/authConfig";
import { UserAgentApplication, AuthError, AuthResponse } from "msal";

export const useGetAccessToken = () => {
  const [accessToken, setAccessToken] = useState<string>("");
  const [accessTokenError, setAccessTokenError] = useState<string>("");

  const msalInstance: UserAgentApplication = new UserAgentApplication(config.msalConfig);

  console.log(`msalInstance: ${msalInstance}`);

  // Power BI REST API call to refresh User Permissions in Power BI
    // Refreshes user permissions and makes sure the user permissions are fully updated
    // https://docs.microsoft.com/rest/api/power-bi/users/refreshuserpermissions
    const tryRefreshUserPermissions = (): void => {
      fetch("https://api.powerbi.com/v1.0/myorg/RefreshUserPermissions", {
          headers: {
              "Authorization": "Bearer " + accessToken
          },
          method: "POST"
      })
      .then(response => {
          if (response.ok) {
              console.log("User permissions refreshed successfully.");
          } else {
              // Too many requests in one hour will cause the API to fail
              if (response.status === 429) {
                  console.error("Permissions refresh will be available in up to an hour.");
              } else {
                  console.error(response);
              }
          }
      })
      .catch(refreshError => {
          console.error("Failure in making API call." + refreshError);
      });
  };

  const successCallback = (response: AuthResponse): void => {
      if(response.tokenType === "id_token") {
        useGetAccessToken();
      } else if (response.tokenType === "access_token") {
          console.log(`successCallbackresponse: ${response}`);
          setAccessToken(response.accessToken);
          tryRefreshUserPermissions();
      } else {
        setAccessTokenError(`Token type is: ${response.tokenType}`);
      }
  };

  const failCallback = (failError: AuthError): void => {
    setAccessTokenError(`Redirect error: ${failError}`);
  };

  msalInstance.handleRedirectCallback(successCallback,failCallback);

  // //check if there is a cached user
  // if (msalInstance.getAccount()) {
    //get access token silently from cached id-token
    msalInstance.acquireTokenSilent(config.loginRequest)
      .then((response:AuthResponse) => {
        console.log(`aquireTokenSilentResponse: ${response}`);
        //get access token from response: response.accessToken
        setAccessToken(response.accessToken);
      })
      .catch((err: AuthError) => {
        //refresh access token silently from cached id-token
        //makes the call to handleredirectcallback
        if(err.name === "InteractionRequiredAuthError") {
          msalInstance.acquireTokenRedirect(config.loginRequest);
        }
        else {
          setAccessTokenError(err.toString());
        }
      });
  // } else {
  //   //user is not logged in or cached, we need to log them in to acquire a token
  //   msalInstance.loginRedirect(config.loginRequest);
  // }
  return {accessToken, accessTokenError};
};
