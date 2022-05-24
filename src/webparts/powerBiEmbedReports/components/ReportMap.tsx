import * as React from 'react';
import styles from './PowerBiEmbedReports.module.scss';
import { PowerBIEmbed } from 'powerbi-client-react';
import { models } from 'powerbi-client';
import { useGetAccessToken } from '../hooks/useGetAccessToken';
import { Carousel, CarouselButtonsDisplay, CarouselButtonsLocation, CarouselIndicatorsDisplay, CarouselIndicatorShape } from '@pnp/spfx-controls-react/lib/Carousel';

export const ReportMap = (props) => {
  const { reports } = props;
  const { accessToken, accessTokenError } = useGetAccessToken();

  reports.forEach(report => {console.log(report.ReportId);});

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
      indicators={true}
      indicatorShape={CarouselIndicatorShape.square}
      indicatorsDisplay={CarouselIndicatorsDisplay.block}
      element={reportsMap}
      interval={null}
    />
  )
  }
  else
  {
    return (
      <div></div>
    )
  }
};
