import * as React from 'react';
import styles from './PowerBiEmbedReports.module.scss';
import { PowerBIEmbed } from 'powerbi-client-react';
import { models } from 'powerbi-client';
import * as config from '../config/authConfig';
import { useGetAccessToken } from '../hooks/useGetAccessToken';

export const App = () => {
  const { accessToken, embedUrl, error } = useGetAccessToken();

  return (
    <div>
      {error &&
        <div>${error}</div>
      }

      {(!error) &&
        <PowerBIEmbed
          embedConfig={{
            type: 'report',   // Supported types: report, dashboard, tile, visual and qna
            id: config.reportId,
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

        // getEmbeddedComponent={(embeddedReport) => {
        //   window.report = embeddedReport;
        // }}
        />
      }
    </div>
  )
};


