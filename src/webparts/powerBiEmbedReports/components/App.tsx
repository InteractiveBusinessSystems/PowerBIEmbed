import * as React from 'react';
import styles from './PowerBiEmbedReports.module.scss';
import { PowerBIEmbed } from 'powerbi-client-react';
import { models } from 'powerbi-client';
import * as config from '../config/authConfig';
import { useGetAccessToken } from '../hooks/useGetAccessToken';
import { useReportsList } from '../hooks/useReportsList';
import { useEffect } from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { ReportMap } from './ReportMap';

export const App = () => {
  const { accessToken, accessTokenError } = useGetAccessToken();
  const { state, getReportsListResults, reportsListDispatch } = useReportsList();
  const { data, reportsListIsLoading, reportsListError } = state;

  useEffect(() => {
    getReportsListResults();
  }, [getReportsListResults]);

  if (reportsListIsLoading) {
  return (
    <Spinner size={SpinnerSize.large} />
  );
  }
  if (accessTokenError) {
    return (
      <div>error: {JSON.stringify(accessTokenError)}</div>
    );
  }
  if (reportsListError) {
    return (
      <div>error: {JSON.stringify(reportsListError)}</div>
    )
  }
  return (
    <ReportMap
      reports={data}
    />
    // <div>
    //   {data.map((report, index) => {if(index === 0) {console.log(report.ReportUrl);}})}
    //   <PowerBIEmbed
    //     embedConfig={{
    //       type: 'report',   // Supported types: report, dashboard, tile, visual and qna
    //       id: config.reportId,
    //       embedUrl: 'https://app.powerbi.com/reportEmbed',
    //       accessToken: accessToken,
    //       tokenType: models.TokenType.Aad,
    //       settings: {
    //         panes: {
    //           filters: {
    //             expanded: false,
    //             visible: true
    //           }
    //         },
    //         // background: models.BackgroundType.Transparent,
    //       }
    //     }}

    //     eventHandlers={
    //       new Map([
    //         ['loaded', function () { console.log('Report loaded'); }],
    //         ['rendered', function () { console.log('Report rendered'); }],
    //         ['error', function (event) { console.log(event.detail); }]
    //       ])
    //     }

    //     cssClassName={styles.embeddedReport}

    //   />
    // </div>
  )
};

