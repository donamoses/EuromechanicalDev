import * as React from 'react';
import styles from './EmecApprovalWp.module.scss';
import { IEmecApprovalWpProps } from './IEmecApprovalWpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, IDropdownOption, Label, Link, MessageBar, MessageBarType, PrimaryButton, TextField } from 'office-ui-fabric-react';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import SimpleReactValidator from 'simple-react-validator';
import { sp, IList, Web } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import * as moment from 'moment';
import ReactHtmlParser, { processNodes, convertNodeToElement, htmlparser2 } from 'react-html-parser';
import * as strings from 'EmecApprovalWpWebPartStrings';
// import { MSGraphClient, HttpClientResponse, IHttpClientOptions, HttpClient } from '@microsoft/sp-http';
import {MSGraphClient, HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';
import replaceString from 'replace-string';
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { useMediaQuery } from 'react-responsive';
import * as _ from 'lodash';
// Alert messages
export interface IMessage {
  isShowMessage: boolean;
  messageType: number;
  message: string;
}
export interface IEmecApprovalWpState {
  requester: any;
  requesterComments: any;
  dccComments: any;
  hideProject: boolean;
  publishOptionKey: string;
  publishOption: string;
  status: string;
  statusKey: string;
  approveDocument: string;

  documentID: string;
  documentName: string;
  revision: any;
  ownerName: string;
  ownerEmail: string;
  dueDate: any;
  requesterName: string;
  requesterEmail: string;
  requestedDate: any;
  requesterComment: string;
  linkToDoc: any;
  reviewerData: any[];
  access: string;
  accessDeniedMsgBar: string;
  hidepublish: boolean;
  statusMessage: IMessage;
  comments: string;
  criticalDocument: any;
  approverName: string;
  cancelConfirmMsg: string;
  confirmDialog: boolean;
  savedisable: boolean;
  taskID: any;
  dccreviewerData: any[];
  revisionLevel: any;
  acceptanceCodearray: any[];
  acceptanceCode: any;
  hideacceptance: boolean;
  externalDocument: any;
  hidetransmittalrevision: boolean;
  publishcheck: any;
  transmittalRevision: any;
  projectName: any;
  projectNumber: any;
}
const Desktop = ({ children }) => {

  const isDesktop = useMediaQuery({ minWidth: 501, maxWidth: 10000 });

  return isDesktop ? children : null;

};

// For mobile view  

const Mobile = ({ children }) => {

  const isMobile = useMediaQuery({ maxWidth: 500 });



  return isMobile ? children : null;



};
export default class EmecApprovalWp extends React.Component<IEmecApprovalWpProps, IEmecApprovalWpState, {}> {
  private validator: SimpleReactValidator;
  public constructor(props: IEmecApprovalWpProps) {
    super(props);
    this.state = {
      publishOptionKey: "",
      requester: "",
      linkToDoc: "",
      requesterComments: "",
      dueDate: "",
      dccComments: "",
      hideProject: true,
      publishOption: "",
      status: "",
      statusKey: "",
      approveDocument: 'none',

      documentID: "",
      documentName: "",
      revision: "",
      ownerName: "",
      ownerEmail: "",
      requesterName: "",
      requesterEmail: "",
      requestedDate: "",
      requesterComment: "",
      reviewerData: [],
      access: "none",
      accessDeniedMsgBar: "none",
      hidepublish: true,
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      comments: "",
      criticalDocument: "",
      approverName: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      savedisable: false,
      taskID: "",
      dccreviewerData: [],
      revisionLevel: "",
      acceptanceCodearray: [],
      acceptanceCode: "",
      hideacceptance: true,
      externalDocument: "",
      hidetransmittalrevision: true,
      transmittalRevision: "",
      publishcheck: "",
      projectName: "",
      projectNumber: ""
    };

    this._queryParamGetting = this._queryParamGetting.bind(this);
    this._accessGroups = this._accessGroups.bind(this);
    this._openRevisionHistory = this._openRevisionHistory.bind(this);
    this._bindApprovalForm = this._bindApprovalForm.bind(this);
    this._project = this._project.bind(this);
    this._drpdwnPublishFormat = this._drpdwnPublishFormat.bind(this);
    this._status = this._status.bind(this);
    this._commentsChange = this._commentsChange.bind(this);
    this._saveAsDraft = this._saveAsDraft.bind(this);
    this._docSave = this._docSave.bind(this);
    this._publish = this._publish.bind(this);
    this._returnDoc = this._returnDoc.bind(this);
    this._sendMail = this._sendMail.bind(this);
    this._onCancel = this._onCancel.bind(this);
    this._acceptanceChanged = this._acceptanceChanged.bind(this);
    this._revisionCoding = this._revisionCoding.bind(this);
    this._publishUpdate = this._publishUpdate.bind(this);
  }
  private workflowHeaderID;
  private documentIndexID;
  private sourceDocumentID;
  private workflowDetailID;
  private reciever;
  private currentEmail;
  private reqWeb = Web(this.props.hubUrl);
  private documentApprovedSuccess;
  private documentSavedAsDraft;
  private documentRejectSuccess;
  private documentReturnSuccess;
  private today;
  private revisionLogId;
  private currentrevision;
  private invalidApprovalLink;
  private invalidUser;
  private status = "";
  private redirectUrl = this.props.siteUrl + this.props.redirectUrl;
  private valid;
  // Validator
  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "Please enter mandatory fields"
      }
    });

  }
  //Page Load
  public async componentDidMount() {
    // Get User Messages
    await this._userMessageSettings();
    //Get Current User
    const user = await sp.web.currentUser.get();
    let userEmail = user.Email;
    this.currentEmail = userEmail;
    //Get Today
    this.today = new Date();
    //Get Parameter from URL
    await this._queryParamGetting();
     //Get Approver
     const HeaderItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.workflowHeaderID).select("Approver/ID,Approver/EMail,DocumentIndexID").expand("Approver").get();
     let approverEmail = HeaderItem.Approver.EMail;
     this.documentIndexID = HeaderItem.DocumentIndexID;
    
    //  if(this.valid == "ok"){
    //   //Get Access
    //   await  this._accessGroups();
    // }
    // else if(this.valid == "Validok"){
       //Check Current User is approver
    if (userEmail == approverEmail) {
      this.setState({ access: "", accessDeniedMsgBar: "none", });
      if (this.props.project) {
        this.setState({ hideProject: false });
        await this._project();
      }
      await this._bindApprovalForm();
    }
    else {
      this.setState({
        access: "none",
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.invalidUser, messageType: 4 }
      });
    }
    // }
   
  
  }
  //Get Parameter from URL
  private _queryParamGetting() {
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let headerid = params.get('hid');
    let detailid = params.get('dtlid');
    if (headerid != "" && headerid != null && detailid != "" && detailid != null) {
      this.workflowHeaderID = parseInt(headerid);
      this.workflowDetailID = parseInt(detailid);
      this.valid = "ok";
    }
    else {
      this.setState({
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.invalidApprovalLink, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(this.redirectUrl);
      }, 10000);
    }
  }
  //Get Access Groups
  private async _accessGroups(){
    let accessGroup=[];
    let ok = "No";
    if(this.props.project){
      accessGroup= await this.reqWeb.getList("/sites/"+this.props.hubsite+"/Lists/"+this.props.PermissionMatrixSettings).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendApprovalWF'").get();
    }
    else{
      accessGroup= await this.reqWeb.getList("/sites/"+this.props.hubsite+"/Lists/"+this.props.PermissionMatrixSettings).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendApprovalWF'").get();
    }

let accessGroupItems:any[]= accessGroup[0].AccessGroups.split(',');
console.log(accessGroupItems);
const documentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).select("DepartmentID").get();
console.log(documentIndexItem);
let deptid = parseInt(documentIndexItem.DepartmentID);
const departmentItem: any = await this.reqWeb.getList("/sites/"+this.props.hubsite+"/Lists/"+this.props.departmentList).items.filter('Title eq '+deptid).select("AccessGroups").get();
let AG =departmentItem[0].AccessGroups;
const accessGroupItem: any = await this.reqWeb.getList("/sites/"+this.props.hubsite+"/Lists/"+this.props.accessGroupDetailsList).items.get();
let accessGroupID;
console.log(accessGroupItem.length);
for (let a = 0;a<accessGroupItem.length;a++){
  if(accessGroupItem[a].Title == AG){
    accessGroupID = accessGroupItem[a].GroupID;
  }
}
const postURL = "https://prod-05.uaecentral.logic.azure.com:443/workflows/60862323b80c44369d5bc091f5490bfa/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QzrPWl7wN5e6k873vy-X9qNeBk0VJojo1M7CzwslVsA";

const requestHeaders: Headers = new Headers();
requestHeaders.append("Content-type", "application/json");
const body: string = JSON.stringify({
  'Groupid':accessGroupID,
  'CurrentUserMail':this.currentEmail
  
});
const postOptions: IHttpClientOptions = {
   headers: requestHeaders,
   body: body
};
let responseText: string = "";
let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
let responseJSON = await response.json();
responseText = JSON.stringify(responseJSON);
 console.log(responseJSON);
if (response.ok) {
  
  console.log(responseJSON.ValidUser);
  if(responseJSON.ValidUser == "Yes"){
    this.valid = "Validok";
  }
  else{
//  this.setState({
//     access: "none",
//     accessDeniedMsgBar: "",
//     statusMessage: { isShowMessage: true, message: this.invalidUser, messageType: 4 }
//   });
//   setTimeout(() => {
//     this.setState({ accessDeniedMsgBar: 'none', });
//     window.location.replace(this.redirectUrl);
//   }, 10000);
  }
 
  
}
else {
}
}
  //Bind Approval Form
  public async _bindApprovalForm() {

    let ApproverId;
    let ApproverName;
    let RequesterName;
    let RequesterEmail;
    let RequestedDate;
    let RequesterComment;
    let DueDate;
    let DocumentID;
    let DocumentName;
    let OwnerName;
    let OwnerEmail;
    let Revision;
    let LinkToDocument;
    let ApproverComment;
    var ReviewerArr: any[] = [];
    let ReviewDate;
    let CriticalDocument;
    let TaskID;
    let Status;
    let PublishOption;

    //Get Header Item
    const HeaderItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.workflowHeaderID).select("Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,SourceDocumentID,DocumentIndexID,RequestedDate,RequesterComment,DueDate,PublishFormat").expand("Requester,Approver").get();
    ApproverId = HeaderItem.Approver.ID;
    ApproverName = HeaderItem.Approver.Title;
    this.sourceDocumentID = HeaderItem.SourceDocumentID;
    this.documentIndexID = HeaderItem.DocumentIndexID;
    RequesterName = HeaderItem.Requester.Title;
    RequesterEmail = HeaderItem.Requester.EMail;
    if(HeaderItem.RequestedDate != null){
    var reqdate = new Date(HeaderItem.RequestedDate);
    RequestedDate = moment(reqdate).format('DD-MMM-YYYY HH:mm');
    }
    RequesterComment = HeaderItem.RequesterComment;
    var duedate = new Date(HeaderItem.DueDate);
    DueDate = moment(duedate).format('DD-MMM-YYYY HH:mm');
    PublishOption = HeaderItem.PublishFormat;
    //Get Document Index
    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).select("DocumentID,DocumentName,Owner/Title,Owner/EMail,Revision,SourceDocument,CriticalDocument").expand("Owner").get();
    console.log(DocumentIndexItem);
    DocumentID = DocumentIndexItem.DocumentID;
    DocumentName = DocumentIndexItem.DocumentName;
    OwnerName = DocumentIndexItem.Owner.Title;
    OwnerEmail = DocumentIndexItem.Owner.EMail;
    Revision = DocumentIndexItem.Revision;
    LinkToDocument = DocumentIndexItem.SourceDocument.Url;
    CriticalDocument = DocumentIndexItem.CriticalDocument;
    //Get Workflow Details
    const DetailItem: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.filter("HeaderID eq " + this.workflowHeaderID).select("ID,Workflow,ResponseDate,ResponsibleComment,ResponseStatus,Responsible/Title,TaskID").expand("Responsible").get();
    for (var k in DetailItem) {
      if (DetailItem[k].Workflow == 'Review') {
        var rewdate = new Date(DetailItem[k].ResponseDate);
        ReviewDate = moment(rewdate).format('DD-MMM-YYYY HH:mm');
        ReviewerArr.push({
          ResponseDate: ReviewDate,
          Reviewer: DetailItem[k].Responsible.Title,
          ResponsibleComment: DetailItem[k].ResponsibleComment
        });
      }
      else if (DetailItem[k].Workflow == 'Approval') {
        ApproverComment = DetailItem[k].ResponsibleComment;
        TaskID = DetailItem[k].TaskID;
        if (DetailItem[k].ResponseStatus == "Published") {
          Status = "Approved";
        }
        else {
          Status = DetailItem[k].ResponseStatus;
        }
        if (DetailItem[k].ResponseStatus != "Under Approval") {
          this.setState({ savedisable: true });
        }
      }

    }
    this.setState({
      documentID: DocumentID,
      documentName: DocumentName,
      linkToDoc: LinkToDocument,
      revision: Revision,
      ownerName: OwnerName,
      ownerEmail: OwnerEmail,
      dueDate: DueDate,
      requesterName: RequesterName,
      requesterEmail: RequesterEmail,
      requestedDate: RequestedDate,
      requesterComment: RequesterComment,
      reviewerData: ReviewerArr,
      comments: ApproverComment,
      criticalDocument: CriticalDocument,
      approverName: ApproverName,
      taskID: TaskID,
      statusKey: Status,
      publishOptionKey: PublishOption

    });
    await this._userMessageSettings();
  }

  //Messages
  private async _userMessageSettings() {
    
    const userMessageSettings: any[] = await this.reqWeb.getList("/sites/"+this.props.hubsite+"/Lists/"+this.props.userMessageSettings).items.select("Title,Message").filter("PageName eq 'Approve'").get();
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title == "ApproveSubmitSuccess") {
        var successmsg = userMessageSettings[i].Message;
        this.documentApprovedSuccess = replaceString(successmsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title == "ApproveDraftSuccess") {
        this.documentSavedAsDraft = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title == "InvalidApprovalLink") {
        this.invalidApprovalLink = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title == "InvalidApproverUser") {
        this.invalidUser = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title == "ApproveRejectSuccess") {
        var rejectmsg = userMessageSettings[i].Message;
        this.documentRejectSuccess = replaceString(rejectmsg, '[DocumentName]', this.state.documentName);
      }
      else if (userMessageSettings[i].Title == "ApproveReturnSuccess") {
        var returnmsg = userMessageSettings[i].Message;
        this.documentReturnSuccess = replaceString(returnmsg, '[DocumentName]', this.state.documentName);
      }
    }

  }
  public async _project() {
    let ReviewDate;
    let DCCReviewerArr: any[] = [];
    let Acceptancearray = [];
    let sorted_Acceptance = [];
    let ProjectName;
    let ProjectNumber;
    const HeaderItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.workflowHeaderID).select("RevisionLevel/Id,RevisionLevel/Title,DocumentControllerId,RevisionCodingId,ApproveInSameRevision,DocumentIndexID,AcceptanceCodeId").expand("RevisionLevel").get();
    let DCC = HeaderItem.DocumentControllerId;
    let RevisionLevel = HeaderItem.RevisionLevel.Title;
    let DocumentIndexId = HeaderItem.DocumentIndexID;
    let AcceptanceCode = HeaderItem.AcceptanceCodeId;
    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(DocumentIndexId).select("ExternalDocument,TransmittalDocument,TransmittalRevision").get();
    let ExternalDocument = DocumentIndexItem.ExternalDocument;
    let TransmittalDocument = DocumentIndexItem.TransmittalDocument;
    let TransmittalRevision = DocumentIndexItem.TransmittalRevision;

    const projectInformation  = await  sp.web.getList(this.props.siteUrl + "/Lists/"+this.props.projectInformationListName).items.get();
    console.log("projectInformation",projectInformation);
    if(projectInformation.length>0){
      for(var k in projectInformation){             
    if(projectInformation[k].Key == "ProjectName"){
      this.setState({
        projectName:projectInformation[k].Title,
      });
    } 
    if(projectInformation[k].Key == "ProjectNumber"){
      this.setState({
      projectNumber:projectInformation[k].Title,
      });
    }
      }
    }
    if (DCC != null) {
      const DetailItem: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.filter("HeaderID eq " + this.workflowHeaderID).select("ID,Workflow,ResponseDate,ResponsibleComment,ResponseStatus,Responsible/Title,TaskID").expand("Responsible").get();
      for (var l in DetailItem) {
        if (DetailItem[l].Workflow == 'DCC Review') {
          var rewdate = new Date(DetailItem[l].ResponseDate.toString()).toLocaleString();
          ReviewDate = moment(rewdate).format('DD-MMM-YYYY HH:mm');
          DCCReviewerArr.push({
            ResponseDate: ReviewDate,
            Reviewer: DetailItem[l].Responsible.Title,
            DCCResponsibleComment: DetailItem[l].ResponsibleComment
          });
        }
      }
    }
    if (ExternalDocument == true) {
      this.setState({ hideacceptance: false });
      const transmittalcodeitems: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.transmittalCodeSettingsList).items.getAll();

      for (let i = 0; i < transmittalcodeitems.length; i++) {
        if (transmittalcodeitems[i].AcceptanceCode == true) {
          let transmittalcodedata = {
            key: transmittalcodeitems[i].ID,
            text: transmittalcodeitems[i].Description
          };
          Acceptancearray.push(transmittalcodedata);
        }
      }
      console.log(Acceptancearray);
      sorted_Acceptance = _.orderBy(Acceptancearray, 'text', ['asc']);

    }
    if (TransmittalDocument == true) {
      this.setState({ hidetransmittalrevision: false });
    }
    this.setState({
      dccreviewerData: DCCReviewerArr,
      revisionLevel: RevisionLevel,
      acceptanceCodearray: sorted_Acceptance,
      externalDocument: ExternalDocument,
      transmittalRevision: TransmittalRevision,
      acceptanceCode: AcceptanceCode
    });
  }
  //Revision History Url
  private _openRevisionHistory = () => {
    window.open(this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.documentIndexID);
  }
  private _onPublishTransmittal = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked == true) {
      this.setState({ publishcheck: isChecked });
      this._revisionCoding();
    }
  }

  //Status Change
  public _status(option: { key: any; text: any }) {
    //console.log(option.key);


    if (option.key == 'Approved') {
      this.setState({ hidepublish: false });
    }
    else {
      this.setState({ hidepublish: true });

    }
    this.setState({ statusKey: option.key, status: option.text, savedisable: false });
  }
  //Publish Change
  public _drpdwnPublishFormat(option: { key: any; text: any }) {
    //console.log(option.key);
    this.setState({ publishOptionKey: option.key, publishOption: option.text, savedisable: false });
  }
  public async _acceptanceChanged(option: { key: any; text: any }) {
    console.log(option.key);
    this.setState({ acceptanceCode: option.key });
  }
  //Comment Change
  public _commentsChange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ comments: comments, savedisable: false });
  }
  //Save as Draft
  public _saveAsDraft = async () => {
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(this.workflowDetailID).update({

      ResponsibleComment: this.state.comments,

    });
    this.setState({
      statusMessage: { isShowMessage: true, message: this.documentSavedAsDraft, messageType: 4 }
    });
  }
  //Data Save
  private _docSave = async () => {
    this.setState({ savedisable: true });

    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLogList).items.filter("WorkflowID eq '" + this.workflowHeaderID + "' and (Workflow eq 'Approve')").get().then(ifyes => {
      this.revisionLogId = ifyes[0].ID;
    });

    if (this.state.hidepublish == false) {
      if (this.validator.fieldValid("publishFormat") && this.validator.fieldValid("status")) {
        this.validator.hideMessages();
        this.setState({ approveDocument: "" });
        setTimeout(() => this.setState({ approveDocument: 'none', }), 3000);
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.workflowHeaderID).update({
          PublishFormat: this.state.publishOption,

        });
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(this.workflowDetailID).update({
          ResponsibleComment: this.state.comments,
          ResponseStatus: "Published",
          ResponseDate: this.today,
        });

        this._publish();
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();

      }
    }
    else {
      if (this.validator.fieldValid("status")) {
        this.validator.hideMessages();
        this.setState({ approveDocument: "" });
        setTimeout(() => this.setState({ approveDocument: 'none' }), 3000);

        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowDetailsList).items.getById(this.workflowDetailID).update({
          ResponsibleComment: this.state.comments,
          ResponseStatus: this.state.status,
          ResponseDate: this.today

        });
        this._returnDoc();

      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }

  }
  
  public _revisionCoding = async () => {

    if (this.props.project) {
      let revision = parseInt(this.state.revision);
      let rev = revision + 1;
      this.currentrevision = rev.toString();
    }
    else {
      let revision = parseInt(this.state.revision);
      let rev = revision + 1;
      this.currentrevision = rev.toString();
    }
  }
  //Document Published
  protected async _publish(){
    this._revisionCoding();
    const postURL = "https://prod-22.uaecentral.logic.azure.com:443/workflows/2fec0e6b5b3642a692a466951503751a/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TXMLhWRMwxiaNOQ9HMGE5C_qsRBqhJ70uQx5ccEUgTE";

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'Status': 'Published',
      'PublishFormat': this.state.publishOption,
      'SourceDocumentID': this.sourceDocumentID,
      'SiteURL': this.props.siteAddress,
      'DocumentName': this.state.documentName,
      'PublishedDocumentLibrary': this.props.publishedDocument,
      'Revision': this.currentrevision,
      'PublishedDate': this.today,
      'SourceDocumentLibrary':this.props.sourceDocumentLibrary
    });
    const postOptions: IHttpClientOptions = {
       headers: requestHeaders,
       body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    responseText = JSON.stringify(responseJSON);
     console.log(responseJSON);
    if (response.ok) {
      
      this._publishUpdate(responseJSON.PublishDocID);
    }
    else {
    }
  }
public async _publishUpdate(publishid){
  let SD =  await sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocument).items.getById(this.sourceDocumentID).get();

  await sp.web.getList(this.props.siteUrl + "/" + this.props.publishedDocument).items.getById(publishid).update({
    DocumentID:this.state.documentID,
    DocumentName:this.state.documentName,
    DocumentIndexId:this.documentIndexID,
    PublishedDate:this.today,
    BusinessUnit:SD.BusinessUnit,
    Category:SD.Category,
    SubCategory:SD.SubCategory,
    ApproverId:SD.ApproverId,
    PublishFormat: this.state.publishOption,
    WorkflowStatus: "Published",
    Revision: this.currentrevision,
    ExpiryLeadPeriod:SD.ExpiryLeadPeriod,
    OwnerId:SD.OwnerId,
    RevisionHistory: {
      "__metadata": { type: "SP.FieldUrlValue" },
      Description: "Revision Log",
      Url: this.props.siteUrl + "/SitePages/RevisionHistory.aspx?did=" + this.documentIndexID
    },
    ReviewersId: { results: SD.ReviewersId },
  });

if (this.state.hideProject == true) {
                await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).update({
                  PublishFormat: this.state.publishOption,
                  WorkflowStatus: "Published",
                  Revision: this.currentrevision
                });
          
                await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.workflowHeaderID).update({
                  ApprovedDate: this.today,
                  WorkflowStatus: "Published",
                  PublishFormat: this.state.publishOption,
                  Revision: this.currentrevision,
                });
          
                await sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocument).items.getById(this.sourceDocumentID).update({
                  PublishFormat: this.state.publishOption,
                  WorkflowStatus: "Published",
                  Revision: this.currentrevision
                });
              }
              else {
                await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).update({
                  PublishFormat: this.state.publishOption,
                  WorkflowStatus: "Published",
                  Revision: this.currentrevision,
                  AcceptanceCodeId: parseInt(this.state.acceptanceCode)
                });
          
                await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.workflowHeaderID).update({
                  ApprovedDate: this.today,
                  WorkflowStatus: "Published",
                  PublishFormat: this.state.publishOption,
                  Revision: this.currentrevision,
                  AcceptanceCodeId: parseInt(this.state.acceptanceCode)
                });
                await sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocument).items.getById(this.sourceDocumentID).update({
                  PublishFormat: this.state.publishOption,
                  WorkflowStatus: "Published",
                  Revision: this.currentrevision,
                  AcceptanceCodeId: parseInt(this.state.acceptanceCode)
                });
              }
              this._sendMail(this.state.ownerEmail, "DocPublish", this.state.ownerName);
              this._sendMail(this.state.requesterEmail, "DocPublish", this.state.requesterName);
              let a = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLogList).items.getById(this.revisionLogId).update({
                Status: "Published"
              });
              await this.reqWeb.getList("/sites/"+this.props.hubsite+"/Lists/"+this.props.workflowTasksList).items.getById(parseInt(this.state.taskID)).delete();
              setTimeout(() => {
                this.setState({ statusMessage: { isShowMessage: true, message: this.documentApprovedSuccess, messageType: 4 } });
                window.location.replace(this.redirectUrl);
              }, 10000);

}
 
  //Document Return
  public _returnDoc = async () => {
    let message;
    let logstatus;
    if (this.state.status == "Rejected" && this.state.hideProject == false) {
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus: this.state.status,
        AcceptanceCodeId: parseInt(this.state.acceptanceCode)
      });

      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.workflowHeaderID).update({
        ApprovedDate: this.today,
        WorkflowStatus: this.state.status,
        AcceptanceCodeId: parseInt(this.state.acceptanceCode)
      });

      await sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocument).items.getById(this.sourceDocumentID).update({
        WorkflowStatus: this.state.status,
        AcceptanceCodeId: parseInt(this.state.acceptanceCode)
      });
    }
    else {
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus: this.state.status
      });

      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.workflowHeaderList).items.getById(this.workflowHeaderID).update({
        ApprovedDate: this.today,
        WorkflowStatus: this.state.status
      });

      await sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocument).items.getById(this.sourceDocumentID).update({
        WorkflowStatus: this.state.status
      });
    }
    if (this.state.status == "Returned with comments") {
      message = this.documentReturnSuccess;
      logstatus = "Approval with Return with comments";
      this._sendMail(this.state.ownerEmail, "DocReturn", this.state.ownerName);
      this._sendMail(this.state.requesterEmail, "DocReturn", this.state.requesterName);

    }
    else {
      message = this.documentRejectSuccess;
      logstatus = "Rejected";
      this._sendMail(this.state.ownerEmail, "DocRejected", this.state.ownerName);
      this._sendMail(this.state.requesterEmail, "DocRejected", this.state.requesterName);

    }

    let a = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLogList).items.getById(this.revisionLogId).update({
      Status: logstatus
    });
    await this.reqWeb.getList("/sites/"+this.props.hubsite+"/Lists/"+this.props.workflowTasksList).items.getById(parseInt(this.state.taskID)).delete();
    setTimeout(() => {
      this.setState({
        statusMessage: { isShowMessage: true, message: message, messageType: 4 }
      });
      window.location.replace(this.redirectUrl);
    }, 10000);
  }
  //Send Mail
  public _sendMail = async (emailuser, type, name) => {

    let formatday = moment(this.today).format('DD/MMM/YYYY');
    let day = formatday.toString();
    let mailSend = "No";
    let Subject;
    let Body;
    console.log(this.state.criticalDocument);

    const notificationPreference: any[] = await this.reqWeb.getList("/sites/"+this.props.hubsite+"/Lists/"+this.props.notificationPreference).items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference").get();
    console.log(notificationPreference[0].Preference);
    if(notificationPreference.length>0){
      if (notificationPreference[0].Preference == "Send all emails") {
        mailSend = "Yes";
      }
      else if (notificationPreference[0].Preference == "Send mail for critical document" && this.state.criticalDocument == true) {
        mailSend = "Yes";
  
      }
      else {
        mailSend = "No";
      } 
    }
    else if(this.state.criticalDocument == true){        
        //console.log("Send mail for critical document");
        mailSend="Yes";         
    } 
    if (mailSend == "Yes") {
      const emailNotification: any[] = await this.reqWeb.getList("/sites/"+this.props.hubsite+"/Lists/"+this.props.emailNotification).items.get();
      console.log(emailNotification);
      for (var k in emailNotification) {
        if (emailNotification[k].Title == type) {
          Subject = emailNotification[k].Subject;
          Body = emailNotification[k].Body;
        }

      }
      let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
      let replaceRequester = replaceString(Body, '[Sir/Madam]', name);
      let replaceDate = replaceString(replaceRequester, '[PublishedDate]', day);
      let replaceApprover = replaceString(replaceDate, '[Approver]', this.state.approverName);
      let replaceBody = replaceString(replaceApprover, '[DocumentName]', this.state.documentName);

      //Create Body for Email  
      let emailPostBody: any = {
        "message": {
          "subject": replacedSubject,
          "body": {
            "contentType": "Text",
            "content": replaceBody

          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": emailuser
              }
            }
          ],
        }
      };

      //Send Email uisng MS Graph  
      this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client
            .api('/me/sendMail')
            .post(emailPostBody);
        });
    }
  }
  //Cancel Document
  private _onCancel = () => {
    this.setState({
      cancelConfirmMsg: "",
      confirmDialog: false,
    });


  }
  //Cancel confirm
  private _confirmYesCancel = () => {
    this.setState({
      statusKey: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    this.validator.hideMessages();
    window.location.replace(this.redirectUrl);
  }
  //Not Cancel
  private _confirmNoCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    //this.validator.hideMessages();
    // window.location.replace(this.RedirectUrl);
  }
  //access denied msgbar close button click
  private _closeButton = () => {
    window.location.replace(this.redirectUrl);
  }

  //For dialog box of cancel
  private _dialogCloseButton = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
  }
  private dialogStyles = { main: { maxWidth: 500 } };
  private dialogContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to cancel?',
    //subText: '<b>Do you want to cancel? </b> ',
  };
  private modalProps = {
    isBlocking: true,
  };


  public render(): React.ReactElement<IEmecApprovalWpProps> {
    const Status: IDropdownOption[] = [
      { key: 'Approved', text: 'Approved' },
      { key: 'Returned with comments', text: 'Returned with comments' },
      { key: 'Rejected', text: 'Rejected' },
    ];
    const PublishOption: IDropdownOption[] = [
      { key: 'PDF', text: 'PDF' },
      { key: 'Native', text: 'Native' },
    ];
    return (
      <div className={styles.emecApprovalWp}>
        <Desktop>
          <div style={{ display: this.state.access }}>
            <div className={styles.alignCenter}> Approval form</div>
            <br></br>
            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>

            <div style={{ display: "flex" }}>
              <div className={styles.width}><Label >Document ID : {this.state.documentID}</Label></div>
              <div><Link onClick={this._openRevisionHistory} underline>Revision History</Link></div>
            </div>

            <div >
              <Label >Document : <a href={this.state.linkToDoc}>{this.state.documentName}</a></Label>
              <div hidden={this.state.hideProject}>
                <div className={styles.flex} >
                  <div className={styles.width}><Label >Project Name : {this.state.projectName}</Label></div>
                  <div ><Label>Project Number : {this.state.projectNumber} </Label></div>
                </div></div>
              <div className={styles.flex}>
                <div className={styles.width}><Label >Revision : {this.state.revision}</Label></div>
                <div hidden={this.state.hideProject}><Label>Revision Level : {this.state.revisionLevel} </Label></div>
              </div>
              <div className={styles.flex}>
                <div className={styles.width}><Label >Owner : {this.state.ownerName} </Label></div>
                <div><Label >Due Date : {this.state.dueDate}</Label></div>
              </div>
              <div className={styles.flex}>
                <div className={styles.width}><Label>Requester : {this.state.requesterName} </Label></div>
                <div><Label >Requested Date : {this.state.requestedDate} </Label></div>
              </div>
              <div className={styles.flex}>
                <div><Label> Requester Comment : </Label>{ReactHtmlParser(this.state.requesterComment)}</div>
              </div>
              <br></br>
              <div hidden={this.state.hideProject} >

                <div style={{ display: (this.state.dccreviewerData.length == 0 ? 'none' : 'block') }}>
                  <table className={styles.tableClass}   >
                    <tr className={styles.tr}>
                      <th className={styles.th}>Document Controller</th>
                      <th className={styles.th}>Document Controller Date</th>
                      <th className={styles.th}>Document Controller Comment</th>
                    </tr>
                    <tbody className={styles.tbody}>
                      {this.state.dccreviewerData.map((item) => {
                        return (<tr className={styles.tr}>
                          <td className={styles.th}>{item.Reviewer}</td>
                          <td className={styles.th}>{moment.utc(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                          <td className={styles.th}>{item.DCCResponsibleComment}</td>
                        </tr>);
                      })
                      }

                    </tbody>
                  </table>
                </div>
              </div>
              <br></br>
              <div style={{ display: (this.state.reviewerData.length == 0 ? 'none' : 'block') }}>
                <table className={styles.tableClass}   >
                  <tr className={styles.tr}>
                    <th className={styles.th}>Reviewer</th>
                    <th className={styles.th}>Review Date</th>
                    <th className={styles.th}>Review Comment</th>
                  </tr>
                  <tbody className={styles.tbody}>
                    {this.state.reviewerData.map((item) => {
                      return (<tr className={styles.tr}>
                        <td className={styles.th}>{item.Reviewer}</td>
                        <td className={styles.th}>{moment.utc(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                        <td className={styles.th}>{item.ResponsibleComment}</td>
                      </tr>);
                    })
                    }

                  </tbody>
                </table>
              </div>
            </div>
            <div >
              <div className={styles.mt}>
                <div hidden={this.state.hideProject}>
                  <div className={styles.flex} >
                    <div className={styles.width}><Label >Transmittal Revision : {this.state.transmittalRevision}</Label></div>
                    <div ><Checkbox label="Publish For Transmittal " onChange={this._onPublishTransmittal} /></div>
                  </div>
                </div>
              </div>
              <div className={styles.mt}>
                <Dropdown
                  placeholder="Select Status"
                  label="Status"
                  options={Status}
                  onChanged={this._status}
                  selectedKey={this.state.statusKey}
                  required />
                <div style={{ color: "#dc3545" }}>{this.validator.message("status", this.state.statusKey, "required")}{" "}</div>
              </div>
              <div className={styles.mt} hidden={this.state.hidepublish}>
                <Dropdown
                  placeholder="Select Option"
                  label="Publish Option"
                  style={{ marginBottom: '10px', backgroundColor: "white" }}
                  options={PublishOption}
                  onChanged={this._drpdwnPublishFormat}
                  selectedKey={this.state.publishOptionKey}
                  required />
                <div style={{ color: "#dc3545" }}>{this.validator.message("publishFormat", this.state.publishOptionKey, "required")}{" "}</div>
              </div>
              <div className={styles.mt} hidden={this.state.hideProject} >
                <div hidden={this.state.hideacceptance}>
                  <Dropdown id="transmittalcode" required={true}
                    placeholder="Select an option"
                    label="Acceptance Code"
                    options={this.state.acceptanceCodearray}
                    onChanged={this._acceptanceChanged}
                    selectedKey={this.state.acceptanceCode}
                  /></div></div>
              <div className={styles.mt}>
                < TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentsChange} multiline required autoAdjustHeight></TextField></div>
              <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div>
              <DialogFooter>

                <div className={styles.rgtalign}>
                  <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
                </div>
                <div className={styles.rgtalign} >
                  <DefaultButton id="b2" className={styles.btn} onClick={this._saveAsDraft}>Save as draft</DefaultButton >

                  <DefaultButton id="b2" className={styles.btn} disabled={this.state.savedisable} onClick={this._docSave}>Submit</DefaultButton >
                  <DefaultButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</DefaultButton >
                </div>
              </DialogFooter>
              {/* {/ Cancel Dialog Box /} */}
              <div style={{ display: this.state.cancelConfirmMsg }}>
                <div>
                  <Dialog
                    hidden={this.state.confirmDialog}
                    dialogContentProps={this.dialogContentProps}
                    onDismiss={this._dialogCloseButton}
                    styles={this.dialogStyles}
                    modalProps={this.modalProps}>
                    <DialogFooter>
                      <PrimaryButton onClick={this._confirmYesCancel} text="Yes" />
                      <DefaultButton onClick={this._confirmNoCancel} text="No" />
                    </DialogFooter>
                  </Dialog>
                </div>
              </div>
              <br />
            </div>
          </div>
          <div style={{ display: this.state.accessDeniedMsgBar }}>

            {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''}
          </div>
        </Desktop>
        <Mobile>
          <div style={{ display: this.state.access }}>
            <div className={styles.alignCenter}> Approval form</div>
            <br></br>
            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>

            <div style={{ display: "flex" }}>
              <Label >Document ID : {this.state.documentID}</Label></div>
            <div>
              <Link onClick={this._openRevisionHistory} underline>
                Revision History
              </Link></div>
            <br></br>
            <div >
              <Label >Document : <a href={this.state.linkToDoc}>{this.state.documentName}</a></Label>

              <div><Label >Revision : {this.state.revision}</Label></div>
              <div hidden={this.state.hideProject}><Label>Revision Level : ABT </Label></div>

              <div ><Label >Owner : {this.state.ownerName} </Label></div>
              <div><Label >Due Date : {this.state.dueDate}</Label></div>

              <div><Label>Requester : {this.state.requesterName} </Label></div>
              <div><Label >Requested Date : {this.state.requestedDate} </Label></div>

              <div><Label> Requester Comment : </Label>{ReactHtmlParser(this.state.requesterComment)}</div>
              <br></br>
              <div hidden={this.state.hideProject} >
                <div className={styles.tableblock} style={{ display: (this.state.dccreviewerData.length == 0 ? 'none' : 'block') }}>
                  <table className={styles.tableClass}   >
                    <tr className={styles.tr}>
                      <th className={styles.th}>Document Controller</th>
                      <th className={styles.th}>Document Controller Date</th>
                      <th className={styles.th}>Document Controller Comment</th>
                    </tr>
                    <tbody className={styles.tbody}>
                      {this.state.dccreviewerData.map((item) => {
                        return (<tr className={styles.tr}>
                          <td className={styles.th}>{item.Reviewer}</td>
                          <td className={styles.th}>{moment.utc(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                          <td className={styles.th}>{item.DCCResponsibleComment}</td>
                        </tr>);
                      })
                      }

                    </tbody>
                  </table>
                </div>
              </div>
              <br></br>
              <div className={styles.tableblock} style={{ display: (this.state.reviewerData.length == 0 ? 'none' : 'block') }}>
                <table className={styles.tableClass}   >
                  <tr className={styles.tr}>
                    <th className={styles.th}>Reviewer</th>
                    <th className={styles.th}>Review Date</th>
                    <th className={styles.th}>Review Comment</th>
                  </tr>
                  <tbody className={styles.tbody}>
                    {this.state.reviewerData.map((item) => {
                      return (<tr className={styles.tr}>
                        <td className={styles.th}>{item.Reviewer}</td>
                        <td className={styles.th}>{moment.utc(item.ResponseDate).format('DD/MM/YYYY, h:mm a')}</td>
                        <td className={styles.th}>{item.ResponsibleComment}</td>
                      </tr>);
                    })
                    }

                  </tbody>
                </table>
              </div>
            </div>
            <div  >
              <div className={styles.mt} hidden={this.state.hideProject}>
                <Checkbox label="Publish For Transmittal " onChange={this._onPublishTransmittal} /></div>
              <div className={styles.mt}>
                <Dropdown
                  placeholder="Select Status"
                  label="Status"
                  options={Status}
                  onChanged={this._status}
                  selectedKey={this.state.statusKey}
                  required /></div>
              <div style={{ color: "#dc3545" }}>{this.validator.message("status", this.state.statusKey, "required")}{" "}</div>
              <div className={styles.mt} hidden={this.state.hidepublish}>
                <Dropdown
                  placeholder="Select Option"
                  label="Publish Option"
                  style={{ marginBottom: '10px', backgroundColor: "white" }}
                  options={PublishOption}
                  onChanged={this._drpdwnPublishFormat}
                  selectedKey={this.state.publishOptionKey}
                  required />
                <div style={{ color: "#dc3545" }}>{this.validator.message("publishFormat", this.state.publishOptionKey, "required")}{" "}</div>
              </div>
              <div className={styles.mt} hidden={this.state.hideProject} >
                <div hidden={this.state.hideacceptance}>
                  <Dropdown id="transmittalcode" required={true}
                    placeholder="Select an option"
                    label="Acceptance Code"
                    options={this.state.acceptanceCodearray}
                    onChanged={this._acceptanceChanged}
                    selectedKey={this.state.acceptanceCode}
                  /></div></div>
              <div className={styles.mt}>
                < TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentsChange} multiline required autoAdjustHeight></TextField></div>
              <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div>
              <DialogFooter>


                <div className={styles.rgtalign}>
                  <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
                </div>
                <div className={styles.rgtalign} >
                  <DefaultButton id="b2" className={styles.btn} onClick={this._saveAsDraft}>Save as draft</DefaultButton >

                  <DefaultButton id="b2" className={styles.btn} disabled={this.state.savedisable} onClick={this._docSave}>Submit</DefaultButton >
                  <DefaultButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</DefaultButton >
                </div>
              </DialogFooter>
              {/* {/ Cancel Dialog Box /} */}
              <div style={{ display: this.state.cancelConfirmMsg }}>
                <div>
                  <Dialog
                    hidden={this.state.confirmDialog}
                    dialogContentProps={this.dialogContentProps}
                    onDismiss={this._dialogCloseButton}
                    styles={this.dialogStyles}
                    modalProps={this.modalProps}>
                    <DialogFooter>
                      <PrimaryButton onClick={this._confirmYesCancel} text="Yes" />
                      <DefaultButton onClick={this._confirmNoCancel} text="No" />
                    </DialogFooter>
                  </Dialog>
                </div>
              </div>
              <br />
            </div>
          </div>
          <div style={{ display: this.state.accessDeniedMsgBar }}>

            {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''}
          </div>
        </Mobile>
      </div>
    );
  }
}
