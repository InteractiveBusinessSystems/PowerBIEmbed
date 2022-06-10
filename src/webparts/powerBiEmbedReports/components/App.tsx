import * as React from 'react';
import { useReportsList } from '../hooks/useReportsList';
import { useEffect, useState } from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react';
import { useGetAccessToken } from '../hooks/useGetAccessToken';
import { useGetReportConfig } from '../hooks/useGetReportConfig';
import { models } from 'powerbi-client';
import { PowerBIEmbed } from 'powerbi-client-react';
import styles from './PowerBiEmbedReports.module.scss';
import {Carousel, CarouselButtonsLocation, CarouselButtonsDisplay, CarouselIndicatorsDisplay, CarouselIndicatorShape} from '@pnp/spfx-controls-react'
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
  const { accessTokenState, getAccessToken } = useGetAccessToken();
  const { accessToken, getAccessTokenIsLoading, getAccessTokenError } = accessTokenState;
  const { reportConfigState, getReportConfig } = useGetReportConfig();
  const { ReportConfig, getReportConfigIsLoading, getReportConfigError } = reportConfigState;
  const [reportId, setReportId]= useState<string>();
  const [embedUrl, setEmbedUrl] = useState<string>();
  const [embedToken, setEmbedToken] = useState<string>();
  let reportsMap;

  useEffect(() => {
    getReportsListResults();
  }, [getReportsListResults]);

  useEffect(() => {
    console.log('reportConfig UseEffect');
    if(isAudienced){
      getReportConfig(aadHttpClient);
      if(!getReportConfigError && !getReportConfigIsLoading){
        setReportId(ReportConfig.ReportId);
        setEmbedToken(ReportConfig.EmbedToken);
        setEmbedUrl(ReportConfig.EmbedUrl);
      }
    }
  },[isAudienced, getReportConfig, setReportId]);



  // useEffect(()=>{
  //   if(isAudienced){
  //     getAccessToken();
  //   }
  // },[isAudienced, getAccessToken]);

  // useEffect(()=> {
    // if (!reportsListIsLoading && !reportsListError) {
    //   reportsMap = reports.map((report) =>
    //     <div>
    //       <a
    //         href={report.ReportUrl}
    //         target="_blank"
    //         className={styles.reportTitle}
    //       >{report.ReportName}</a>
    //       <PowerBIEmbed
    //         embedConfig={{
    //           type: 'report',   // Supported types: report, dashboard, tile, visual and qna
    //           id: report.ReportId,
    //           embedUrl: 'https://app.powerbi.com/reportEmbed',
    //           accessToken: accessToken,
    //           tokenType: models.TokenType.Embed,
    //           settings: {
    //             panes: {
    //               filters: {
    //                 expanded: false,
    //                 visible: true
    //               }
    //             },
    //             // background: models.BackgroundType.Transparent,
    //           }
    //         }}

    //         eventHandlers={
    //           new Map([
    //             ['loaded', function () { console.log('Report loaded'); }],
    //             ['rendered', function () { console.log('Report rendered'); }],
    //             ['error', function (event) { console.log(event.detail); }]
    //           ])
    //         }

    //         cssClassName={styles.embeddedReport}

    //       />
    //     </div>
    //   );
    // }
  // },[reports]);


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
      </div>
    )
  }
  else {
    return (
      <PowerBIEmbed
        embedConfig={{
          type: 'report',   // Supported types: report, dashboard, tile, visual and qna
          id: reportId,
          embedUrl: embedUrl,
          accessToken: embedToken,
          tokenType: models.TokenType.Embed,
          settings: {
            panes: {
              filters: {
                expanded: false,
                visible: true
              }
            },
            // background: models.BackgroundType.Transparent,
          }
        }}

        eventHandlers={
          new Map([
            ['loaded', function () { console.log('Report loaded'); }],
            ['rendered', function () { console.log('Report rendered'); }],
            ['error', function (event) { console.log(event.detail); }]
          ])
        }

        cssClassName={styles.embeddedReport}

      />

        //  <Carousel
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
    )
  }
};

