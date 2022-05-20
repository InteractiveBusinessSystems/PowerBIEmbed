// This didn't work, but I'm keeping it around just in case
import * as React from 'react';
import styles from './PowerBiEmbedReports.module.scss';
import { PowerBIEmbed } from 'powerbi-client-react';
import { models } from 'powerbi-client';
import { useGetAccessToken } from '../hooks/useGetAccessToken';
import { useEffect, useState } from 'react';
import { Carousel, CarouselButtonsDisplay, CarouselButtonsLocation, CarouselIndicatorsDisplay, CarouselIndicatorShape } from '@pnp/spfx-controls-react/lib/Carousel';

export const ReportMap = (props) => {
  const { reports } = props;
  const { accessToken, accessTokenError } = useGetAccessToken();
  const [reportUrl, setReportUrl] = useState<string>("");
  const [reportId, setReportId] = useState<string>("");


  const reportsMap = reports.map((report) =>
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
              embedUrl: 'https://app.powerbi.com/reportEmbed',
              accessToken: accessToken,
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
        </div>
      );

  if(reportsMap){
  return (
    <Carousel
      buttonsLocation={CarouselButtonsLocation.bottom}
      buttonsDisplay={CarouselButtonsDisplay.hidden}
      contentContainerStyles={styles.carouselContent}
      containerButtonsStyles={styles.carouselButtonsContainer}
      indicatorShape={CarouselIndicatorShape.square}
      indicatorsDisplay={CarouselIndicatorsDisplay.block}
      element={reportsMap}
    />
    // <div>
    //   {reportsMap}
    // </div>
  )
  }
  else
  {
    return (
      <div></div>
    )
  }

  // useEffect(() => {
  //   if (reports) {
  //     reports.map((report, index) => {
  //       if (index === 0) {
  //         setReportUrl(report.ReportUrl);
  //         setReportId(report.ReportId);
  //       }
  //     });
  //   }
  // }, [reports]);

  // if (reportUrl !== "") {
  //   return (
  //     <div>
  //       <PowerBIEmbed
  //         embedConfig={{
  //           type: 'report',   // Supported types: report, dashboard, tile, visual and qna
  //           id: reportId,
  //           embedUrl: 'https://app.powerbi.com/reportEmbed',
  //           accessToken: accessToken,
  //           tokenType: models.TokenType.Aad,
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
  //   )
  // }
  // else {
  //   return (
  //     <div></div>
  //   )
  // }

};
