import { ReactText } from "react";

export interface IReportsList {
  ReportName: string;
  WorkspaceId: string;
  ReportId: string;
  ReportSectionId: string;
  ViewerType: string;
  UsersWhoCanView: [];
  Id: number;
}
