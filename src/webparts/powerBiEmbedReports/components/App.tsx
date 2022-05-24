import * as React from 'react';
import { useReportsList } from '../hooks/useReportsList';
import { useEffect } from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { ReportMap } from './ReportMap';

export const App = (props) => {
  const isAudienced = props.isAudienced;
  const { state, getReportsListResults } = useReportsList();
  const { data, reportsListIsLoading, reportsListError } = state;

  useEffect(() => {
    getReportsListResults();
  }, [getReportsListResults]);

  if (reportsListIsLoading) {
  return (
    <Spinner size={SpinnerSize.large} />
  );
  }
  if (reportsListError) {
    return (
      <div>error: {JSON.stringify(reportsListError)}</div>
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
    />
  )
};

