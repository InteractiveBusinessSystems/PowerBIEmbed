import * as React from 'react';
import styles from './PowerBiEmbedReports.module.scss';
import { IPowerBiEmbedReportsProps } from './IPowerBiEmbedReportsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { App } from './App';


export default class PowerBiEmbedReports extends React.Component<IPowerBiEmbedReportsProps, {}> {
  public render(): React.ReactElement<IPowerBiEmbedReportsProps> {
    console.log (this.props.isAudienced);
    return (
      <div className={styles.powerBiEmbedReports}>
        <div className={this.props.isAudienced? styles.container: styles.containerNoAudience}>
        {/* <div className={styles.container}> */}
          {/* <p className={styles.description}>{escape(this.props.description)}</p> */}
          <App
            isAudienced={this.props.isAudienced}
          />
        </div>
      </div>
    );
  }
}
