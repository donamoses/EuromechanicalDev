import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IEmecInboundSubContractorProps {
  description: string;
  context: WebPartContext;
  redirectUrl:string;
  siteUrl:string;
  projectInformationListName:string;
  revisionLevelList:string;
  transmittalCodeSettings:string;
  hubUrl:string;
  hubsite:string;
  companyList:string;
  transmittalOutlookLibrary:string;
  documentIndexList:string;
}
