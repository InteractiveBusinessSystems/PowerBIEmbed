import * as React from 'react';
import { useReportsList } from '../hooks/useReportsList';
import { useEffect, useState } from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { useGetAccessToken } from '../hooks/useGetAccessToken';

export interface IAppProps {
  isAudienced: boolean;
}

export const App = (props:IAppProps) => {
  const {isAudienced} = props;
  const { state, getReportsListResults } = useReportsList();
  const { reports, reportsListIsLoading, reportsListError } = state;
  const { accessTokenState, getAccessToken } = useGetAccessToken();
  const { accessToken, getAccessTokenIsLoading, getAccessTokenError } = accessTokenState;
  const [reportId, setReportId]= useState<string>();

  useEffect(() => {
    getReportsListResults();
  }, [getReportsListResults]);

  useEffect(()=>{
    if(isAudienced){
      getAccessToken();
    }
  },[getAccessToken]);

  useEffect(()=> {
    if (!reportsListIsLoading && !reportsListError) {
      console.log(reports);
      reports.forEach((report, index) => {
        if (index === 0) {
          setReportId(report.ReportId);
        }
      });
    }
  },[reports]);

// const reportsMap = reports.map((report) =>
  //       <div>
  //         <a
  //           href={report.ReportUrl}
  //           target="_blank"
  //           className={styles.reportTitle}
  //         >{report.ReportName}</a>
  //         <PowerBIEmbed
  //           embedConfig={{
  //             type: 'report',   // Supported types: report, dashboard, tile, visual and qna
  //             id: report.ReportId,
  //             embedUrl: 'https://app.powerbi.com/reportEmbed',
  //             accessToken: accessToken,
  //             tokenType: models.TokenType.Aad,
  //             settings: {
  //               panes: {
  //                 filters: {
  //                   expanded: false,
  //                   visible: true
  //                 }
  //               },
  //               // background: models.BackgroundType.Transparent,
  //             }
  //           }}

  //           eventHandlers={
  //             new Map([
  //               ['loaded', function () { console.log('Report loaded'); }],
  //               ['rendered', function () { console.log('Report rendered'); }],
  //               ['error', function (event) { console.log(event.detail); }]
  //             ])
  //           }

  //           cssClassName={styles.embeddedReport}

  //         />
  //       </div>
  //     );

  // if(reportsMap){
  // return (
  //   <Carousel
  //     buttonsLocation={CarouselButtonsLocation.bottom}
  //     buttonsDisplay={CarouselButtonsDisplay.hidden}
  //     contentContainerStyles={styles.carouselContent}
  //     containerButtonsStyles={styles.carouselButtonsContainer}
  //     indicators={true}
  //     indicatorShape={CarouselIndicatorShape.square}
  //     indicatorsDisplay={CarouselIndicatorsDisplay.block}
  //     element={reportsMap}
  //     interval={null}
  //   />
  // )
  // }


  if (reportsListIsLoading) {
    return (
      <Spinner size={SpinnerSize.large} />
    );
  }
  // if (getAccessTokenIsLoading) {
  //   return (
  //     <Spinner size={SpinnerSize.large} />
  //   );
  // }

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
  if (!isAudienced) {
    return (
      <div>
      </div>
    )
  }
  else {
    return (
      <div>
        <h1>Power BI Embed Web Part</h1>
        <p>Reports should go here!!</p>
        <p>reportId: {reportId}</p>
        <p>accessToken: {accessToken}</p>
      </div>
      // <PowerBIEmbed
      //   embedConfig={{
      //     type: 'report',   // Supported types: report, dashboard, tile, visual and qna
      //     id: reportId,
      //     embedUrl: 'https://app.powerbi.com/reportEmbed',
      //     accessToken: accessToken,
      //     tokenType: models.TokenType.Aad,
      //     settings: {
      //       panes: {
      //         filters: {
      //           expanded: false,
      //           visible: true
      //         }
      //       },
      //       // background: models.BackgroundType.Transparent,
      //     }
      //   }}

      //   eventHandlers={
      //     new Map([
      //       ['loaded', function () { console.log('Report loaded'); }],
      //       ['rendered', function () { console.log('Report rendered'); }],
      //       ['error', function (event) { console.log(event.detail); }]
      //     ])
      //   }

      //   cssClassName={styles.embeddedReport}

      // />
    )

  }
};

