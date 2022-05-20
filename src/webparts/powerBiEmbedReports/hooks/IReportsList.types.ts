import { ReactText } from "react";

export interface IReportsList {
  ReportName: string;
  WorkspaceId: string;
  ReportId: string;
  ReportSectionId: string;
  ReportUrl: string;
  ViewerType: string;
  UsersWhoCanView: [];
  Id: number;
}
