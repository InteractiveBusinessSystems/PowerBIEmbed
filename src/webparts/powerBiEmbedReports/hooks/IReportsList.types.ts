import { ReactText } from "react";

export interface IReportsList {
  ReportName: string;
  DataSetsId?: string;
  WorkspaceId: string;
  ReportId: string;
  ReportUrl: string;
  ViewerType: string;
  UsersWhoCanView: [];
  Id: number | undefined;
  EmbedToken?: string;
  EmbedUrl?: string;
  AccessToken?: string;
}
