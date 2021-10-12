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
import { MSGraphClient, HttpClientResponse, IHttpClientOptions, HttpClient } from '@microsoft/sp-http';
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
  Requester: any;
  RequesterComments: any;
  DCCComments: any;
  hideproject: boolean;
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
  RequesterName: string;
  requesterEmail: string;
  requestedDate: any;
  RequesterComment: string;
  linkToDoc: any;
  reviewerData: any[];
  access: string;
  accessDeniedMsgBar: string;
  hidepublish: boolean;
  statusMessage: IMessage;
  comments: string;
  CriticalDocument: any;
  ApproverName: string;
  cancelConfirmMsg: string;
  confirmDialog: boolean;
  savedisable: boolean;
  TaskID: any;
  dccreviewerData: any[];
  RevisionLevel: any;
  AcceptanceCodearray: any[];
  AcceptanceCode: any;
  hideacceptance: boolean;
  ExternalDocument: any;
  hidetransmittalrevision: boolean;
  publishcheck: any;
  TransmittalRevision: any;
  ProjectName: any;
  ProjectNumber: any;
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
      Requester: "",
      linkToDoc: "",
      RequesterComments: "",
      dueDate: "",
      DCCComments: "",
      hideproject: true,
      publishOption: "",
      status: "",
      statusKey: "",
      approveDocument: 'none',

      documentID: "",
      documentName: "",
      revision: "",
      ownerName: "",
      ownerEmail: "",
      RequesterName: "",
      requesterEmail: "",
      requestedDate: "",
      RequesterComment: "",
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
      CriticalDocument: "",
      ApproverName: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      savedisable: false,
      TaskID: "",
      dccreviewerData: [],
      RevisionLevel: "",
      AcceptanceCodearray: [],
      AcceptanceCode: "",
      hideacceptance: true,
      ExternalDocument: "",
      hidetransmittalrevision: true,
      TransmittalRevision: "",
      publishcheck: "",
      ProjectName: "",
      ProjectNumber: ""
    };

    this.queryParamGetting = this.queryParamGetting.bind(this);
    this.accessGroups = this.accessGroups.bind(this);
    this._openRevisionHistory = this._openRevisionHistory.bind(this);
    this.BindApprovalForm = this.BindApprovalForm.bind(this);
    this._drpdwnPublishFormat = this._drpdwnPublishFormat.bind(this);
    this._status = this._status.bind(this);
    this._commentschange = this._commentschange.bind(this);
    this._SaveasDraft = this._SaveasDraft.bind(this);
    this._docSave = this._docSave.bind(this);
    this._publish = this._publish.bind(this);
    this._returndoc = this._returndoc.bind(this);
    this._sendmail = this._sendmail.bind(this);
    this._onCancel = this._onCancel.bind(this);
    this.AcceptanceChanged = this.AcceptanceChanged.bind(this);
    this._revisioncoding = this._revisioncoding.bind(this);
  }
  private workflowHeaderID;
  private documentIndexID;
  private sourceDocumentID;
  private workflowDetailID;
  private Reciever;
  private CurrentEmail;
  private reqWeb = Web(this.props.hubUrl);
  private documentApprovedSuccess;
  private documentSavedAsDraft;
  private documentRejectSuccess;
  private documentReturnSuccess;
  private today;
  private RevisionLogId;
  private currentrevision;
  private InvalidApprovalLink;
  private InvalidUser;
  private Status = "";
  private RedirectUrl = this.props.siteUrl + this.props.RedirectUrl;
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
    await this.userMessageSettings();
    //Get Current User
    const user = await sp.web.currentUser.get();
    let userEmail = user.Email;
    this.CurrentEmail = userEmail;
    //Get Today
    this.today = new Date();
    //Get Parameter from URL
    this.queryParamGetting();
     //Get Approver
     const HeaderItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.getById(this.workflowHeaderID).select("Approver/ID,Approver/EMail,DocumentIndexID").expand("Approver").get();
     let ApproverEmail = HeaderItem.Approver.EMail;
     this.documentIndexID = HeaderItem.DocumentIndexID;
    
    //Get Access
    // this.accessGroups();
   
    //Check Current User is approver
    if (userEmail == ApproverEmail) {
      this.setState({ access: "", accessDeniedMsgBar: "none", });
      if (this.props.project) {
        this.setState({ hideproject: false });
        await this.project();
      }
      await this.BindApprovalForm();
    }
    else {
      this.setState({
        access: "none",
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.InvalidUser, messageType: 4 }
      });
    }
  }
  //Get Parameter from URL
  private queryParamGetting() {
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let headerid = params.get('hid');
    let detailid = params.get('dtlid');
    if (headerid != "" && headerid != null && detailid != "" && detailid != null) {
      this.workflowHeaderID = parseInt(headerid);
      this.workflowDetailID = parseInt(detailid);
    }
    else {
      this.setState({
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.InvalidApprovalLink, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(this.RedirectUrl);
      }, 5000);
    }
  }
  //Get Access Groups
  private async accessGroups(){
    let AccessGroup=[];
    let ok = "No";
    if(this.props.project){
      AccessGroup= await this.reqWeb.lists.getByTitle(this.props.AccessGroups).items.select("AccessGroups,AccessFields").filter("Title eq 'Project_SendApprovalWF'").get();
    }
    else{
      AccessGroup= await this.reqWeb.lists.getByTitle(this.props.AccessGroups).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendApprovalWF'").get();
    }

let AccessGroupItems:any[]= AccessGroup[0].AccessGroups.split(',');
console.log(AccessGroupItems);
const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).select("DepartmentID").get();
console.log(DocumentIndexItem);
let deptid = parseInt(DocumentIndexItem.DepartmentID);
const DepartmentItem: any = await this.reqWeb.lists.getByTitle(this.props.DepartmentList).items.getById(deptid).select("AccessGroups").get();
console.log(DepartmentItem.AccessGroups);
var result = AccessGroupItems.indexOf(DepartmentItem.AccessGroups);
console.log(result);
const groups = await this.reqWeb.siteGroups();
console.log(groups);
  //User in HO Group
  try {
    let grp1: any[] = await this.reqWeb.siteGroups.getByName(AccessGroupItems[result]).users();
    for (let i = 0; i < grp1.length; i++) {
        if (this.CurrentEmail == grp1[i].Email) {
            ok = "Yes"
        }
      }
  if(ok != "Yes"){
    this.setState({
      accessDeniedMsgBar: "",
      statusMessage: { isShowMessage: true, message: this.InvalidUser, messageType: 1 },
    });
    setTimeout(() => {
      this.setState({ accessDeniedMsgBar: 'none', });
      window.location.replace(this.RedirectUrl);
    }, 5000);
}

}
catch { }

  }
  //Bind Approval Form
  public async BindApprovalForm() {

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
    const HeaderItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.getById(this.workflowHeaderID).select("Requester/ID,Requester/Title,Requester/EMail,Approver/ID,Approver/Title,SourceDocumentID,DocumentIndexID,RequestedDate,RequesterComment,DueDate,PublishFormat").expand("Requester,Approver").get();
    ApproverId = HeaderItem.Approver.ID;
    ApproverName = HeaderItem.Approver.Title;
    this.sourceDocumentID = HeaderItem.SourceDocumentID;
    this.documentIndexID = HeaderItem.DocumentIndexID;
    RequesterName = HeaderItem.Requester.Title;
    RequesterEmail = HeaderItem.Requester.EMail;
    // var reqdate = new Date(HeaderItem.RequestedDate.toString()).toLocaleString();
    // RequestedDate = moment(reqdate).format('DD-MMM-YYYY HH:mm');
    RequesterComment = HeaderItem.RequesterComment;
    var duedate = new Date(HeaderItem.DueDate.toString()).toLocaleString();
    DueDate = moment(duedate).format('DD-MMM-YYYY HH:mm');
    PublishOption = HeaderItem.PublishFormat;
    //Get Document Index
    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).select("DocumentID,DocumentName,Owner/Title,Owner/EMail,Revision,SourceDocument,CriticalDocument").expand("Owner").get();
    console.log(DocumentIndexItem);
    DocumentID = DocumentIndexItem.DocumentID;
    DocumentName = DocumentIndexItem.DocumentName;
    OwnerName = DocumentIndexItem.Owner.Title;
    OwnerEmail = DocumentIndexItem.Owner.EMail;
    Revision = DocumentIndexItem.Revision;
    LinkToDocument = DocumentIndexItem.SourceDocument.Url;
    CriticalDocument = DocumentIndexItem.CriticalDocument;
    //Get Workflow Details
    const DetailItem: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.filter("HeaderID eq " + this.workflowHeaderID).select("ID,Workflow,ResponseDate,ResponsibleComment,ResponseStatus,Responsible/Title,TaskID").expand("Responsible").get();
    for (var k in DetailItem) {
      if (DetailItem[k].Workflow == 'Review') {
        var rewdate = new Date(DetailItem[k].ResponseDate.toString()).toLocaleString();
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
      RequesterName: RequesterName,
      requesterEmail: RequesterEmail,
      requestedDate: RequestedDate,
      RequesterComment: RequesterComment,
      reviewerData: ReviewerArr,
      comments: ApproverComment,
      CriticalDocument: CriticalDocument,
      ApproverName: ApproverName,
      TaskID: TaskID,
      statusKey: Status,
      publishOptionKey: PublishOption

    });
    await this.userMessageSettings();
  }

  //Messages
  private async userMessageSettings() {
    const userMessageSettings: any[] = await this.reqWeb.lists.getByTitle(this.props.userMessageSettings).items.select("Title,Message").filter("PageName eq 'Approve'").get();
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
        this.InvalidApprovalLink = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title == "InvalidApproverUser") {
        this.InvalidUser = userMessageSettings[i].Message;
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
  public async project() {
    let ReviewDate;
    let DCCReviewerArr: any[] = [];
    let Acceptancearray = [];
    let sorted_Acceptance = [];
    let ProjectName;
    let ProjectNumber;
    const HeaderItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.getById(this.workflowHeaderID).select("RevisionLevel/Id,RevisionLevel/Title,DocumentControllerId,RevisionCodingId,ApproveInSameRevision,DocumentIndexID,AcceptanceCode/ID").expand("RevisionLevel,AcceptanceCode").get();
    let DCC = HeaderItem.DocumentControllerId;
    let RevisionLevel = HeaderItem.RevisionLevel.Title;
    let DocumentIndexId = HeaderItem.DocumentIndexID;
    let AcceptanceCode = HeaderItem.AcceptanceCode.ID;
    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(DocumentIndexId).select("ExternalDocument,TransmittalDocument,TransmittalRevision").get();
    let ExternalDocument = DocumentIndexItem.ExternalDocument;
    let TransmittalDocument = DocumentIndexItem.TransmittalDocument;
    let TransmittalRevision = DocumentIndexItem.TransmittalRevision;

    const ProjectInformation: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/ProjectInformation").items.get();
    ProjectName = ProjectInformation[0].Title;
    ProjectNumber = ProjectInformation[1].Title;
    // for(let pro=0;pro<=ProjectInformation.length;pro++){
    //   if(ProjectInformation[pro].Key == "ProjectName"){
    //     ProjectName = ProjectInformation[0].Title;
    //   }
    //   else if(ProjectInformation[pro].Key == "ProjectNumber"){
    //     ProjectNumber = ProjectInformation[1].Title;
    //   }
    // }
    if (DCC != null) {
      const DetailItem: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.filter("HeaderID eq " + this.workflowHeaderID).select("ID,Workflow,ResponseDate,ResponsibleComment,ResponseStatus,Responsible/Title,TaskID").expand("Responsible").get();
      for (var k in DetailItem) {
        if (DetailItem[k].Workflow == 'DCC Review') {
          var rewdate = new Date(DetailItem[k].ResponseDate.toString()).toLocaleString();
          ReviewDate = moment(rewdate).format('DD-MMM-YYYY HH:mm');
          DCCReviewerArr.push({
            ResponseDate: ReviewDate,
            Reviewer: DetailItem[k].Responsible.Title,
            DCCResponsibleComment: DetailItem[k].ResponsibleComment
          });
        }
      }
    }
    if (ExternalDocument == true) {
      this.setState({ hideacceptance: false });
      const transmittalcodeitems: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.TransmittalCodeSettingsList).items.getAll();

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
      RevisionLevel: RevisionLevel,
      AcceptanceCodearray: sorted_Acceptance,
      ExternalDocument: ExternalDocument,
      TransmittalRevision: TransmittalRevision,
      ProjectName: ProjectName,
      ProjectNumber: ProjectNumber,
      AcceptanceCode: AcceptanceCode
    });
  }
  //Revision History Url
  private _openRevisionHistory = () => {
    window.open(this.props.siteUrl + "/SitePages/RevisionHistory.aspx?ID=" + this.documentIndexID);
  }
  private _onPublishTransmittal = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked == true) {
      this.setState({ publishcheck: isChecked });
      this._revisioncoding();
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
  public async AcceptanceChanged(option: { key: any; text: any }) {
    console.log(option.key);
    this.setState({ AcceptanceCode: option.key });
  }
  //Comment Change
  public _commentschange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ comments: comments, savedisable: false });
  }
  //Save as Draft
  public _SaveasDraft = async () => {
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(this.workflowDetailID).update({

      ResponsibleComment: this.state.comments,

    });
    this.setState({
      statusMessage: { isShowMessage: true, message: this.documentSavedAsDraft, messageType: 4 }
    });
  }
  //Data Save
  private _docSave = async () => {
    this.setState({ savedisable: true });

    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.filter("WorkflowID eq '" + this.workflowHeaderID + "' and (Workflow eq 'Approve')").get().then(ifyes => {
      this.RevisionLogId = ifyes[0].ID;
    });

    if (this.state.hidepublish == false) {
      if (this.validator.fieldValid("publishFormat") && this.validator.fieldValid("status")) {
        this.validator.hideMessages();
        this.setState({ approveDocument: "" });
        setTimeout(() => this.setState({ approveDocument: 'none', }), 3000);
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.getById(this.workflowHeaderID).update({
          PublishFormat: this.state.publishOption,

        });
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(this.workflowDetailID).update({
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

        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(this.workflowDetailID).update({
          ResponsibleComment: this.state.comments,
          ResponseStatus: this.state.status,
          ResponseDate: this.today

        });
        this._returndoc();

      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }

  }
  //Document Published
  public _publish = async () => {
    this._revisioncoding();
    const postURL = "https://prod-22.uaecentral.logic.azure.com:443/workflows/2fec0e6b5b3642a692a466951503751a/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TXMLhWRMwxiaNOQ9HMGE5C_qsRBqhJ70uQx5ccEUgTE";

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'Status': 'Published',
      'PublishFormat': this.state.publishOption,
      'SourceDocumentID': this.sourceDocumentID,
      'SiteURL': this.props.siteUrl,
      'DocumentName': this.state.documentName,
      'PublishedDocumentLibrary': this.props.PublishedDocument,
      'Revision': this.currentrevision,
      'PublishedDate': this.today
    });
    const postOptions: IHttpClientOptions = {
       headers: requestHeaders,
       body: body
    };
    let responseText: string = "";
    this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions).then((response: HttpClientResponse) => {
        response.json().then((responseJSON: JSON) => {
            console.log(response);
            responseText = JSON.stringify(responseJSON);
            console.log(responseJSON);
            if (response.ok) {
  }
   else {}
})
.catch((responsee: any) => {
let errMsg: string = `WARNING - error when calling URL ${postURL}. Error = ${responsee.message}`;
 console.log(errMsg);
});
});

  }
//   public _publish = async () => {
//     this._revisioncoding();
//     if (this.state.publishOption == "PDF") {

//       //****Get MS Flow & its response*****/

//       const postURL = "https://prod-22.uaecentral.logic.azure.com:443/workflows/2fec0e6b5b3642a692a466951503751a/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=TXMLhWRMwxiaNOQ9HMGE5C_qsRBqhJ70uQx5ccEUgTE";

//       const body: string = JSON.stringify({
//         'Status': 'Published',
//         'PublishFormat': this.state.publishOption,
//         'SourceDocumentID': this.sourceDocumentID,
//         'SiteURL': this.props.siteUrl,
//         'DocumentName': this.state.documentName,
//         'PublishedDocumentLibrary': this.props.PublishedDocument,
//         'Revision': this.currentrevision,
//         'PublishedDate': this.today

//       });

//       const requestHeaders: Headers = new Headers();
//       requestHeaders.append('Content-type', 'application/json');

//       const httpClientOptions: IHttpClientOptions = {
//         body: body,
//         headers: requestHeaders
//       };
// let responseText:string = "";

//       return this.props.context.httpClient.post(
//         postURL,
//         HttpClient.configurations.v1,
//         httpClientOptions)
//         .then((response: HttpClientResponse): Promise<HttpClientResponse> => {
//           return response.json();
        
//         });

//     }

//     if (this.state.hideproject == true) {
//       await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
//         PublishFormat: this.state.publishOption,
//         WorkflowStatus: "Published",
//         Revision: this.currentrevision
//       });

//       await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.getById(this.workflowHeaderID).update({
//         ApprovedDate: this.today,
//         WorkflowStatus: "Published",
//         PublishFormat: this.state.publishOption,
//         Revision: this.currentrevision,
//       });

//       await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocument).items.getById(this.sourceDocumentID).update({
//         PublishFormat: this.state.publishOption,
//         WorkflowStatus: "Published",
//         Revision: this.currentrevision
//       });
//     }
//     else {
//       await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
//         PublishFormat: this.state.publishOption,
//         WorkflowStatus: "Published",
//         Revision: this.currentrevision,
//         AcceptanceCodeId: parseInt(this.state.AcceptanceCode)
//       });

//       await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.getById(this.workflowHeaderID).update({
//         ApprovedDate: this.today,
//         WorkflowStatus: "Published",
//         PublishFormat: this.state.publishOption,
//         Revision: this.currentrevision,
//         AcceptanceCodeId: parseInt(this.state.AcceptanceCode)
//       });
//       await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocument).items.getById(this.sourceDocumentID).update({
//         PublishFormat: this.state.publishOption,
//         WorkflowStatus: "Published",
//         Revision: this.currentrevision,
//         AcceptanceCodeId: parseInt(this.state.AcceptanceCode)
//       });
//     }
//     this._sendmail(this.state.ownerEmail, "DocPublish", this.state.ownerName);
//     this._sendmail(this.state.requesterEmail, "DocPublish", this.state.RequesterName);
//     let a = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.getById(this.RevisionLogId).update({
//       Status: "Published"
//     });
//     await this.reqWeb.lists.getByTitle(this.props.WorkflowTasksList).items.getById(parseInt(this.state.TaskID)).delete();
//     setTimeout(() => {
//       this.setState({ statusMessage: { isShowMessage: true, message: this.documentApprovedSuccess, messageType: 4 } });
//       window.location.replace(this.RedirectUrl);
//     }, 5000);

//   }
  public _revisioncoding = async () => {

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
  //Document Return
  public _returndoc = async () => {
    let message;
    let logstatus;
    if (this.state.status == "Rejected" && this.state.hideproject == false) {
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus: this.state.status,
        AcceptanceCodeId: parseInt(this.state.AcceptanceCode)
      });

      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.getById(this.workflowHeaderID).update({
        ApprovedDate: this.today,
        WorkflowStatus: this.state.status,
        AcceptanceCodeId: parseInt(this.state.AcceptanceCode)
      });

      await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocument).items.getById(this.sourceDocumentID).update({
        WorkflowStatus: this.state.status,
        AcceptanceCodeId: parseInt(this.state.AcceptanceCode)
      });
    }
    else {
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus: this.state.status
      });

      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.getById(this.workflowHeaderID).update({
        ApprovedDate: this.today,
        WorkflowStatus: this.state.status
      });

      await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocument).items.getById(this.sourceDocumentID).update({
        WorkflowStatus: this.state.status
      });
    }
    if (this.state.status == "Returned with comments") {
      message = this.documentReturnSuccess;
      logstatus = "Approval with Return with comments";
      this._sendmail(this.state.ownerEmail, "DocReturn", this.state.ownerName);
      this._sendmail(this.state.requesterEmail, "DocReturn", this.state.RequesterName);

    }
    else {
      message = this.documentRejectSuccess;
      logstatus = "Rejected";
      this._sendmail(this.state.ownerEmail, "DocRejected", this.state.ownerName);
      this._sendmail(this.state.requesterEmail, "DocRejected", this.state.RequesterName);

    }

    let a = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.getById(this.RevisionLogId).update({
      Status: logstatus
    });
    await this.reqWeb.lists.getByTitle(this.props.WorkflowTasksList).items.getById(parseInt(this.state.TaskID)).delete();
    setTimeout(() => {
      this.setState({
        statusMessage: { isShowMessage: true, message: message, messageType: 4 }
      });
      window.location.replace(this.RedirectUrl);
    }, 5000);
  }
  //Send Mail
  public _sendmail = async (emailuser, type, name) => {

    let formatday = moment(this.today).format('DD/MMM/YYYY');
    let day = formatday.toString();
    let mailSend = "No";
    let Subject;
    let Body;
    console.log(this.state.CriticalDocument);

    const notificationPreference: any[] = await this.reqWeb.lists.getByTitle(this.props.notificationPreference).items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference").get();
    console.log(notificationPreference[0].Preference);
    if(notificationPreference.length>0){
      if (notificationPreference[0].Preference == "Send all emails") {
        mailSend = "Yes";
      }
      else if (notificationPreference[0].Preference == "Send mail for critical document" && this.state.CriticalDocument == true) {
        mailSend = "Yes";
  
      }
      else {
        mailSend = "No";
      } 
    }
    else if(this.state.CriticalDocument == true){        
        //console.log("Send mail for critical document");
        mailSend="Yes";         
    } 
    if (mailSend == "Yes") {
      const emailNotification: any[] = await this.reqWeb.lists.getByTitle(this.props.emailNotification).items.get();
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
      let replaceApprover = replaceString(replaceDate, '[Approver]', this.state.ApproverName);
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
    window.location.replace(this.RedirectUrl);
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
    window.location.replace(this.RedirectUrl);
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
              <div hidden={this.state.hideproject}>
                <div className={styles.flex} >
                  <div className={styles.width}><Label >Project Name : {this.state.ProjectName}</Label></div>
                  <div ><Label>Project Number : {this.state.ProjectNumber} </Label></div>
                </div></div>
              <div className={styles.flex}>
                <div className={styles.width}><Label >Revision : {this.state.revision}</Label></div>
                <div hidden={this.state.hideproject}><Label>Revision Level : {this.state.RevisionLevel} </Label></div>
              </div>
              <div className={styles.flex}>
                <div className={styles.width}><Label >Owner : {this.state.ownerName} </Label></div>
                <div><Label >Due Date : {this.state.dueDate}</Label></div>
              </div>
              <div className={styles.flex}>
                <div className={styles.width}><Label>Requester : {this.state.RequesterName} </Label></div>
                <div><Label >Requested Date : {this.state.requestedDate} </Label></div>
              </div>
              <div className={styles.flex}>
                <div><Label> Requester Comment : </Label>{ReactHtmlParser(this.state.RequesterComment)}</div>
              </div>
              <br></br>
              <div hidden={this.state.hideproject} >

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
                <div hidden={this.state.hideproject}>
                  <div className={styles.flex} >
                    <div className={styles.width}><Label >Transmittal Revision : {this.state.TransmittalRevision}</Label></div>
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
              <div className={styles.mt} hidden={this.state.hideproject} >
                <div hidden={this.state.hideacceptance}>
                  <Dropdown id="transmittalcode" required={true}
                    placeholder="Select an option"
                    label="Acceptance Code"
                    options={this.state.AcceptanceCodearray}
                    onChanged={this.AcceptanceChanged}
                    selectedKey={this.state.AcceptanceCode}
                  /></div></div>
              <div className={styles.mt}>
                < TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentschange} multiline required autoAdjustHeight></TextField></div>
              <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div>
              <DialogFooter>

                <div className={styles.rgtalign}>
                  <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
                </div>
                <div className={styles.rgtalign} >
                  <DefaultButton id="b2" className={styles.btn} onClick={this._SaveasDraft}>Save as draft</DefaultButton >

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
              <div hidden={this.state.hideproject}><Label>Revision Level : ABT </Label></div>

              <div ><Label >Owner : {this.state.ownerName} </Label></div>
              <div><Label >Due Date : {this.state.dueDate}</Label></div>

              <div><Label>Requester : {this.state.RequesterName} </Label></div>
              <div><Label >Requested Date : {this.state.requestedDate} </Label></div>

              <div><Label> Requester Comment : </Label>{ReactHtmlParser(this.state.RequesterComment)}</div>
              <br></br>
              <div hidden={this.state.hideproject} >
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
              <div className={styles.mt} hidden={this.state.hideproject}>
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
              <div className={styles.mt} hidden={this.state.hideproject} >
                <div hidden={this.state.hideacceptance}>
                  <Dropdown id="transmittalcode" required={true}
                    placeholder="Select an option"
                    label="Acceptance Code"
                    options={this.state.AcceptanceCodearray}
                    onChanged={this.AcceptanceChanged}
                    selectedKey={this.state.AcceptanceCode}
                  /></div></div>
              <div className={styles.mt}>
                < TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentschange} multiline required autoAdjustHeight></TextField></div>
              <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div>
              <DialogFooter>


                <div className={styles.rgtalign}>
                  <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
                </div>
                <div className={styles.rgtalign} >
                  <DefaultButton id="b2" className={styles.btn} onClick={this._SaveasDraft}>Save as draft</DefaultButton >

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
