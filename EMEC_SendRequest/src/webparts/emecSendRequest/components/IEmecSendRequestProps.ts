import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IEmecSendRequestProps {
  context: WebPartContext;
  siteUrl:string;
  RedirectUrl:string;
  hubUrl:string;
  userMessageSettings:string;
  DocumentIndexList:string;
  project: string;
  notificationPreference:string;
  emailNotification:string;
  WorkflowHeaderList:string;
  TransmittalCodeSettingsList:string;
  WorkflowDetailsList:string;
  DocumentRevisionLogList:string;
  WorkflowTasksList:string;
  SourceDocumentLibrary:string;
  RevisionLevelList:string;
  TaskDelegationSettings:string;
  RevisionHistoryPage:string;
  DocumentApprovalPage:string;
  DocumentReviewPage:string;
  AccessGroups:string;
  DepartmentList:string;
  AccessGroupDetailsList:string;
  hubsite:string;
  projectInformationListName:string;
}
