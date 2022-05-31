import * as React from 'react';
import styles from './PowerBiEmbedReports.module.scss';
import { useCheckMemberGroupsAgainstGroups } from '../hooks/useCheckMemberGroupsAgainstGroups';
import { App } from './App';
import { useEffect } from 'react';
import { IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls';

export interface IAppWrapperProps {
  // description?: string;
  groups?: IPropertyFieldGroupOrPerson[];
  userGroups?: string[];
  accessToken?: string;
  accessTokenError?: string;
}

export const AppWrapper = (props:IAppWrapperProps) => {
  const audiencedGroups = props.groups;
  const userGroups = props.userGroups;
  const accessToken = props.accessToken;
  const accessTokenError = props.accessTokenError;
  const { state, checkMemberGroupsAgainstGroups } = useCheckMemberGroupsAgainstGroups();
  const { data, checkMemberGroupsIsLoading, checkMemberGroupsError} = state;
  const isAudienced = data;

  useEffect(() => {
    if(userGroups){
      checkMemberGroupsAgainstGroups(audiencedGroups,userGroups);
    }
  },[checkMemberGroupsAgainstGroups, audiencedGroups, userGroups]);


  if(checkMemberGroupsError){
    return (
      <div>{JSON.stringify(checkMemberGroupsError)}</div>
    )
  }
  if (isAudienced && !checkMemberGroupsIsLoading && !checkMemberGroupsError) {
    return (
      <div className={styles.powerBiEmbedReports}>
        <div className={styles.container}>
          {/* <p className={styles.description}>{escape(this.props.description)}</p> */}
          <App
            isAudienced={isAudienced}
            accessToken={accessToken}
            accessTokenError={accessTokenError}
          />
        </div>
      </div>
    );
  }
  else {
    return (
      <div className={styles.powerBiEmbedReports}>
        <div className={styles.containerNoAudience}>
        </div>
      </div>
    );
  }
}

