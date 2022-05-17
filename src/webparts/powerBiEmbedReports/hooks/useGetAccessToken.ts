import * as React from "react";
import { useState } from "react";
import * as config from "../config/authConfig";
import { UserAgentApplication, AuthError, AuthResponse } from "msal";

export const useGetAccessToken = () => {
  const [accessToken, setAccessToken] = useState<string>("");
  const [embedUrl, setEmbedUrl] = useState<string>("");
  const [userName, setUsername] = useState<string>("");
  const [error, setError] = useState<string>("");

  const msalInstance: UserAgentApplication = new UserAgentApplication(config.msalConfig);

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

  // Power BI REST API call to get the embed URL of the report
  const getembedUrl = (): void  => {

      fetch("https://api.powerbi.com/v1.0/myorg/groups/" + config.workspaceId + "/reports/" + config.reportId + "/ReportSection" + config.reportsectionId, {
          headers: {
              "Authorization": "Bearer " + accessToken
          },
          method: "GET"
      })
          .then(response => {
              console.log(response);
              response.json()
                  .then(body => {
                      // Successful response
                      if (response.ok) {
                          console.log(`EmbedUrl: ${body["embedUrl"]}`);
                          setEmbedUrl(body["embedUrl"]);
                          // setAccessToken(accessToken);
                      }
                      // If error message is available
                      else {
                          setError("Error " + response.status + ": " + body.error.code);
                      }

                  })
                  .catch(embedResponse => {
                      setError("Error " + embedResponse.status + ":  An error has occurred");
                  });
          })
          .catch(embedError => {

              // Error in making the API call
              setError(embedError);
          });
  };

  const successCallback = (response: AuthResponse): void => {
      if(response.tokenType === "id_token") {
        useGetAccessToken();
      } else if (response.tokenType === "access_token") {
          setAccessToken(response.accessToken);
          setUsername(response.account.name);

          tryRefreshUserPermissions();
          // getembedUrl();
      } else {
        setError(`Token type is: ${response.tokenType}`);
      }
  };

  const failCallback = (failError: AuthError): void => {
    setError(`Redirect error: ${failError}`);
  };

  msalInstance.handleRedirectCallback(successCallback,failCallback);

  //check if there is a cached user
  if (msalInstance.getAccount()) {
    //get access token silently from cached id-token
    msalInstance.acquireTokenSilent(config.loginRequest)
      .then((response:AuthResponse) => {
        //get access token from response: response.accessToken
        setAccessToken(response.accessToken);
        setUsername(response.account.name);
        // getembedUrl();
      })
      .catch((err: AuthError) => {
        //refresh access token silently from cached id-token
        //makes the call to handleredirectcallback
        if(err.name === "InteractionRequiredAuthError") {
          msalInstance.acquireTokenRedirect(config.loginRequest);
        }
        else {
          setError(err.toString());
        }
      });
  } else {
    //user is not logged in or cached, we need to log them in to acquire a token
    msalInstance.loginRedirect(config.loginRequest);
  }
  console.log(`AccessToken: ${accessToken}`);
  console.log(`embedUrl: ${embedUrl}`);
  console.log(`error: ${error}`);
  console.log(`userName: ${userName}`);
  return {accessToken, embedUrl, userName, error};
};
