// This didn't work, but I'm keeping it around just in case
import * as React from 'react';
import styles from './PowerBiEmbedReports.module.scss';
import { PowerBIEmbed } from 'powerbi-client-react';
import { models } from 'powerbi-client';
import { useGetAccessToken } from '../hooks/useGetAccessToken';

export const ReportMap = (props) => {
  const { reports } = props;
  const { accessToken, accessTokenError } = useGetAccessToken();

  return (
    <div>
      {reports.forEach ((report) => {
        console.log(report);
        console.log(report.ReportId);
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
        })}
    </div>
  );
};
