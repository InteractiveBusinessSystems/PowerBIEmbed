import * as React from 'react';
import { useReportsList } from '../hooks/useReportsList';
import { useEffect } from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { ReportMap } from './ReportMap';
import { useGetAccessToken } from '../hooks/useGetAccessToken';

export interface IAppProps {
  isAudienced: boolean;
}

export const App = (props:IAppProps) => {
  const {isAudienced} = props;
  const { state, getReportsListResults } = useReportsList();
  const { data, reportsListIsLoading, reportsListError } = state;
  const { accessTokenState, getAccessToken } = useGetAccessToken();
  const { accessToken, getAccessTokenIsLoading, getAccessTokenError } = accessTokenState;

  useEffect(() => {
    getReportsListResults();
  }, [getReportsListResults]);

  useEffect(()=>{
    if(isAudienced){
      getAccessToken();
    }
  },[getAccessToken]);

  if (reportsListIsLoading) {
  return (
    <Spinner size={SpinnerSize.large} />
  );
  }
  // if (getAccessTokenIsLoading) {
  //   return (
  //     <Spinner size={SpinnerSize.large} />
  //   );
  //   }
  if (reportsListError) {
    return (
      <div>error: {JSON.stringify(reportsListError)}</div>
    )
  }
  if (getAccessTokenError) {
    return (
      <div>error: {JSON.stringify(getAccessTokenError)}</div>
    )
  }
  if(!isAudienced){
    return (
      <div></div>
    )
  }
  return (
    <ReportMap
      reports={data}
      accessToken={accessToken}
    />
  )
};

