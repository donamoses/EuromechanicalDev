import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IEmecApprovalWpProps {
  description: string;
  project: string;
  context: WebPartContext;
  RedirectUrl: string;
  siteUrl:string;
  hubUrl:string;
  notificationPreference:string;
  emailNotification:string;
  userMessageSettings:string;
  WorkflowHeaderList:string;
  DocumentIndexList:string;
  TransmittalCodeSettingsList:string;
  WorkflowDetailsList:string;
  SourceDocument:string;
  PublishedDocument:string;
  DocumentRevisionLogList:string;
  WorkflowTasksList:string;
  AccessGroups:string;
  DepartmentList:string;
}

