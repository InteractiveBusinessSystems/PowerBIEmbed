import * as React from 'react';
import { useReportsList } from '../hooks/useReportsList';
import { useEffect } from 'react';
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
      reportsMap = ReportConfig.map((report) =>
        <div>
          <a
            href={report.ReportUrl}
            target="_blank"
            className={styles.reportTitle}
          >{report.ReportName}</a>
          <PowerBIEmbed
            embedConfig={{
              type: 'report',   // Supported types: report, dashboard, tile, visual and qna
              id: report.ReportId,
              embedUrl: `https://app.powerbi.com/reportEmbed?reportId=${report.ReportId}&groupId=${report.WorkspaceId}`,
              accessToken: report.accessToken,
              tokenType: models.TokenType.Aad,
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
          {/* <PowerBIEmbed
            embedConfig={{
              type: 'report',   // Supported types: report, dashboard, tile, visual and qna
              id: report.ReportId,
              embedUrl: report.EmbedUrl,
              accessToken: report.EmbedToken,
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

          /> */}
        </div>
      );
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
    )
  }
};

