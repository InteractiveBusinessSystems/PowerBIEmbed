import * as React from 'react';
import { useReportsList } from '../hooks/useReportsList';
import { useEffect, useState } from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { useGetReportConfig } from '../hooks/useGetReportConfig';
import { models } from 'powerbi-client';
import { PowerBIEmbed } from 'powerbi-client-react';
import styles from './PowerBiEmbedReports.module.scss';
import {Carousel, CarouselButtonsLocation, CarouselButtonsDisplay, CarouselIndicatorsDisplay, CarouselIndicatorShape, ICarouselImageProps} from '@pnp/spfx-controls-react'
import { AadHttpClientFactory } from '@microsoft/sp-http';

export interface IAppProps {
  isAudienced: boolean;
  aadHttpClient: AadHttpClientFactory;
}

export const App = (props:IAppProps) => {
  const isAudienced = props.isAudienced;
  const aadHttpClient = props.aadHttpClient;
  const { state, getReportsListResults } = useReportsList();
  const { reports, reportsListIsLoading, reportsListError } = state;
  const { reportConfigState, getReportConfig } = useGetReportConfig();
  const { ReportConfig, getReportConfigIsLoading, getReportConfigError } = reportConfigState;
  let reportsMap: JSX.Element | JSX.Element[] | ICarouselImageProps[] = [];

  useEffect(() => {
    getReportsListResults();
  }, [getReportsListResults]);

  useEffect(() => {
    if(isAudienced){
      if(reports !== null && !reportsListError && !reportsListIsLoading){
        getReportConfig(aadHttpClient, reports);
      }
    }
  },[isAudienced, reportsListError, reportsListIsLoading, getReportConfig]);

  if (!getReportConfigIsLoading && !getReportConfigError) {
    reportsMap = [];
    // reportsMap = ReportConfig.map((report) =>
    //   <div>
    //     <a
    //       href={report.ReportUrl}
    //       target="_blank"
    //       className={styles.reportTitle}
    //     >{report.ReportName}</a>
    //     <PowerBIEmbed
    //       embedConfig={{
    //         type: 'report',   // Supported types: report, dashboard, tile, visual and qna
    //         id: report.ReportId,
    //         embedUrl: report.EmbedUrl,
    //         accessToken: report.AccessToken,
    //         tokenType: models.TokenType.Aad,
    //         settings: {
    //           panes: {
    //             filters: {
    //               expanded: false,
    //               visible: true
    //             }
    //           },
    //           // background: models.BackgroundType.Transparent,
    //         }
    //       }}

    //       eventHandlers={
    //         new Map([
    //           ['loaded', function () { console.log('Report loaded'); }],
    //           ['rendered', function () { console.log('Report rendered'); }],
    //           ['error', function (event) { console.log(event.detail); }]
    //         ])
    //       }

    //       cssClassName={styles.embeddedReport}
    //     />
    //   </div>
    // );
  }

  if (reportsListIsLoading) {
    return (
      <Spinner size={SpinnerSize.large} />
    );
  }
  if (getReportConfigIsLoading) {
    return (
      <Spinner size={SpinnerSize.large} />
    );
  }
  if (reportsListError) {
    return (
      <div>error: {JSON.stringify(reportsListError)}</div>
    )
  }
  if (getReportConfigError) {
    return (
      <div>error: {JSON.stringify(getReportConfigError)}</div>
    )
  }
  if (!isAudienced) {
    return (
      <div>
        <p>You need to set up your web part. Please edit the web part and select a security group.</p>
      </div>
    )
  }
  else {
    return (
         <Carousel
            buttonsLocation={CarouselButtonsLocation.bottom}
            buttonsDisplay={CarouselButtonsDisplay.hidden}
            contentContainerStyles={styles.carouselContent}
            containerButtonsStyles={styles.carouselButtonsContainer}
            indicators={true}
            indicatorShape={CarouselIndicatorShape.square}
            indicatorsDisplay={CarouselIndicatorsDisplay.block}
            element={reportsMap}
            interval={null}
          />
        // <PowerBIEmbed
        //     embedConfig={{
        //       type: 'report',   // Supported types: report, dashboard, tile, visual and qna
        // id: '79505149-73b0-45d3-98f9-e49b8a2328e2',
        // embedUrl: 'https://app.powerbi.com/reportEmbed?reportId=79505149-73b0-45d3-98f9-e49b8a2328e2&groupId=7af23086-163b-4747-bd1c-977d1830d59b&w=2&config=eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9XQUJJLVVTLU5PUlRILUNFTlRSQUwtcmVkaXJlY3QuYW5hbHlzaXMud2luZG93cy5uZXQiLCJlbWJlZEZlYXR1cmVzIjp7Im1vZGVybkVtYmVkIjp0cnVlLCJ1c2FnZU1ldHJpY3NWTmV4dCI6dHJ1ZSwic2tpcFF1ZXJ5RGF0YVNhYVNFbWJlZCI6dHJ1ZSwic2tpcFF1ZXJ5RGF0YVBhYVNFbWJlZCI6dHJ1ZSwic2tpcFF1ZXJ5RGF0YUV4cG9ydFRvIjp0cnVlfX0%3d',
        // accessToken: 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvNGVjNTU0OTMtNmIxYy00NTY1LWE4NjgtMmFlOTQwODgyYzgyLyIsImlhdCI6MTY2NTA2MjY3OCwibmJmIjoxNjY1MDYyNjc4LCJleHAiOjE2NjUwNjc2MDksImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJBVlFBcS84VEFBQUFYZFpCOCtxYXoxbnpJYVJ6M3VUd1VDUVBqQ3VuNHloeWIweGEyQzhiSjBxbWpPVjUyN1V0c0VUWm5OWGQyUU9wa1lRRlczamxsWkxhZU1Qdy9HOHRJWE9FUi9OQ2VISkxtRU5SL1k3ekk2cz0iLCJhbXIiOlsicHdkIiwibWZhIl0sImFwcGlkIjoiMTcwYWY1NTYtZDI2Yy00MGIzLTlhOTYtMzYxY2UxMWQ2ODNkIiwiYXBwaWRhY3IiOiIwIiwiZmFtaWx5X25hbWUiOiJEYXJyb2NoIiwiZ2l2ZW5fbmFtZSI6IlNoZXJ5bCIsImlwYWRkciI6IjEwNy4xMjYuODEuMTAxIiwibmFtZSI6IlNoZXJ5bCBEYXJyb2NoIiwib2lkIjoiMjNhN2M3ZDgtMjMzMy00ZWExLWE0NDEtY2Y3OGU4NzE2YjQyIiwicHVpZCI6IjEwMDNCRkZEQUFENkExRjIiLCJyaCI6IjAuQVFnQWsxVEZUaHhyWlVXb2FDcnBRSWdzZ2drQUFBQUFBQUFBd0FBQUFBQUFBQUFJQU9nLiIsInNjcCI6IkRhdGFzZXQuUmVhZC5BbGwgUmVwb3J0LlJlYWQuQWxsIFdvcmtzcGFjZS5SZWFkLkFsbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6IlpqWW1nVlkxSnNjZmxGS0tHR2o3NlZUSXM1MkYzVkNYQ3RaX1oxMmIzVEkiLCJ0aWQiOiI0ZWM1NTQ5My02YjFjLTQ1NjUtYTg2OC0yYWU5NDA4ODJjODIiLCJ1bmlxdWVfbmFtZSI6IlNEYXJyb2NoQGliczM2NS5jb20iLCJ1cG4iOiJTRGFycm9jaEBpYnMzNjUuY29tIiwidXRpIjoibVk3Z3Z6VWN2RTZwZFh3S0pFNXNBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il19.DmLzajnnvos_sYeOOsZTcF7rCQVYHVw0POe1DZfNANZXeILF-wNWW4Dp1mTtiWOV7z0O8qHtmaE1GRc1JSweDiV6ww5QnTrEL_LQPruKu_7meeDlvD-SFHHSe4QSqnwm5hKwZmTa19kuaDZ54ZBvwOR2A5YETfakOp_HL_muk7U5A7VhnxbBhII1462FIG-WURk5bpjKOsqg9ejBDCWhbN_7ypNsGdwM1XvCrFRoiovOGQMVcUHTJql5LW3J1vefXT1newsmmidTSoem8iCwBFsQxqIDwMqRDKQ2QdpnHgHvNiTEcJpIjKaCyfvD3rQIgm6iZNQ0PHp4RHs1NP-7Lg',
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
    )
  }
};

