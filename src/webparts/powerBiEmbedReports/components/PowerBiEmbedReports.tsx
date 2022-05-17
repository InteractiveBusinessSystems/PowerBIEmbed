import * as React from 'react';
import styles from './PowerBiEmbedReports.module.scss';
import { IPowerBiEmbedReportsProps } from './IPowerBiEmbedReportsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { App } from './App';


export default class PowerBiEmbedReports extends React.Component<IPowerBiEmbedReportsProps, {}> {
  public render(): React.ReactElement<IPowerBiEmbedReportsProps> {
    return (
      <div className={styles.powerBiEmbedReports}>
        <div className={styles.container}>
          {/* <p className={styles.description}>{escape(this.props.description)}</p> */}
          <App
            context={this.props.context}
          />
        </div>
      </div>
    );
  }
}
