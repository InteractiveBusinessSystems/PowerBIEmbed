import * as React from 'react';
import { IPowerBiEmbedReportsProps } from './IPowerBiEmbedReportsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { AppWrapper } from './AppWrapper';



export default class PowerBiEmbedReports extends React.Component<IPowerBiEmbedReportsProps, {}> {

  public render(): React.ReactElement<IPowerBiEmbedReportsProps> {
    return (
      <AppWrapper
          groups={this.props.groups}
          userGroups={this.props.userGroups}
          aadHttpClient={this.props.aadHttpClient}
      />
    );

  }
}
