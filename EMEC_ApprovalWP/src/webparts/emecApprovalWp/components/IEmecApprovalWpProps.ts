import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IEmecApprovalWpProps {
  description: string;
  project: string;
  context: WebPartContext;
  redirectUrl: string;
  siteUrl:string;
  hubUrl:string;
  notificationPreference:string;
  emailNotification:string;
  userMessageSettings:string;
  workflowHeaderList:string;
  documentIndexList:string;
  transmittalCodeSettingsList:string;
  workflowDetailsList:string;
  sourceDocument:string;
  publishedDocument:string;
  documentRevisionLogList:string;
  workflowTasksList:string;
  PermissionMatrixSettings:string;
  departmentList:string;
  sourceDocumentLibrary:string;
  siteAddress:string;
  accessGroupDetailsList:string;
  hubsite:string;
  projectInformationListName:string;
}

