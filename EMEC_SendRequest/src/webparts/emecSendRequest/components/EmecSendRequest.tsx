import * as React from 'react';
import styles from './EmecSendRequest.module.scss';
import { IEmecSendRequestProps } from './IEmecSendRequestProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, Label, Link, MessageBar, PrimaryButton, TextField } from 'office-ui-fabric-react';
import SimpleReactValidator from 'simple-react-validator';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp, IList, Web } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import * as _ from 'lodash';
import * as moment from 'moment';
import { useMediaQuery } from 'react-responsive';
import { MSGraphClient, HttpClientResponse, IHttpClientOptions, HttpClient } from '@microsoft/sp-http';
import replaceString from 'replace-string';
export interface IMessage {
  isShowMessage: boolean;
  messageType: number;
  message: string;
}
export interface IEmecSendRequestState {
  statusMessage: IMessage;
  documentID: string;
  linkToDoc: string;
  documentName: string;
  revision: any;
  ownerName: string;
  currentUser: string;
  hideProject: boolean;
  revisionLevel: any[];
  revisionLevelvalue: any;
  dcc: any;
  reviewer: any;
  dueDate: any;
  approver: any;
  comments: any;
  cancelConfirmMsg: string;
  confirmDialog: boolean;
  saveDisable: boolean;
  requestSend: string;
  statusKey: string;
  access: any;
  accessDeniedMsgBar: any;
  reviewers: any[];
  currentUserRewviewer: any[];
  ownerId: any;
  delegatedToId: any;
  delegateToIdInSubSite: any;
  delegateForIdInSubSite: any;
  reviewerEmail: any;
  reviewerName: any;
  delegatedFromId: any;
  detailIdForReviewer: any;
  approverEmail: any;
  approverName: any;
  hubSiteUserId: any;
  detailIdForApprover: any;
  criticalDocument: any;
  dccReviewerName: any;
  dccReviewerEmail: any;
  dccReviewer: any;
  revisionLevelArray: any[];
  revisionCoding: any;
  projectName: any;
  projectNumber: any;
  acceptanceCodeId: any;
  transmittalRevision: any;

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
export default class EmecSendRequest extends React.Component<IEmecSendRequestProps, IEmecSendRequestState, {}> {
  private validator: SimpleReactValidator;
  private reqWeb = Web(this.props.hubUrl);
  private documentIndexID;
  private invalidUser;
  private currentEmail;
  private currentId;
  private today;
  private time;
  private workflowStatus;
  private sourceDocumentID;
  private newheaderid;
  private newDetailItemID;
  private DccReview;
  private UnderApproval;
  private UnderReview;
  private redirectUrl = this.props.siteUrl + this.props.RedirectUrl;
  private invalidSendRequestLink;
  private getSelectedReviewers = [];
  private valid;

  public constructor(props: IEmecSendRequestProps) {
    super(props);
    this.state = {

      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      documentID: "",
      linkToDoc: "",
      documentName: "",
      revision: "",
      ownerName: "",
      currentUser: "",
      hideProject: true,
      revisionLevel: [],
      revisionLevelvalue: "",
      dcc: "",
      reviewer: "",
      dueDate: "",
      approver: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      saveDisable: false,
      requestSend: 'none',
      statusKey: "",
      access: "none",
      accessDeniedMsgBar: "none",
      reviewers: [],
      ownerId: "",
      delegatedToId: "",
      delegateToIdInSubSite: "",
      delegateForIdInSubSite: "",
      reviewerEmail: "",
      reviewerName: "",
      delegatedFromId: "",
      detailIdForReviewer: "",
      approverEmail: "",
      approverName: "",
      hubSiteUserId: 0,
      detailIdForApprover: "",
      criticalDocument: "",
      dccReviewerName: "",
      dccReviewerEmail: "",
      dccReviewer: "",
      revisionLevelArray: [],
      revisionCoding: "",
      currentUserRewviewer: [],
      projectName: "",
      projectNumber: "",
      acceptanceCodeId: "",
      transmittalRevision: ""

    };
    this.componentDidMount = this.componentDidMount.bind(this);
    this._userMessageSettings = this._userMessageSettings.bind(this);
    this._queryParamGetting = this._queryParamGetting.bind(this);
    this._accessGroups = this._accessGroups.bind(this);
    this._checkWorkflowStatus = this._checkWorkflowStatus.bind(this);
    this._openRevisionHistory = this._openRevisionHistory.bind(this);
    this._bindSendRequestForm = this._bindSendRequestForm.bind(this);
    this._project = this._project.bind(this);
    this._revisionLevelChanged = this._revisionLevelChanged.bind(this);
    this._dccReviewerChange = this._dccReviewerChange.bind(this);
    this._reviewerChange = this._reviewerChange.bind(this);
    this._approverChange = this._approverChange.bind(this);
    this._submitSendRequest = this._submitSendRequest.bind(this);
    this._dccReview = this._dccReview.bind(this);
    this._underApprove = this._underApprove.bind(this);
    this._underReview = this._underReview.bind(this);
    this._underProjectApprove = this._underProjectApprove.bind(this);
    this._underProjectReview = this._underProjectReview.bind(this);
  }
  // Validator
  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "This field is mandatory"
      }
    });

  }
  //Page Load
  public async componentDidMount() {
    // Get User Messages
    await this._userMessageSettings();
    //Get Parameter from URL
    this._queryParamGetting();

    if (this.props.project) {
      this.setState({ hideProject: false });
    }
    //Get Current User
    const user = await sp.web.currentUser.get();
    this.currentEmail = user.Email;
    this.currentId = user.Id;
    let currentuserrewviewer = [];
    currentuserrewviewer.push(this.currentId);
    //Get Today
    this.today = new Date();
    this.setState({
      currentUser: user.Title,
      currentUserRewviewer: currentuserrewviewer
    });
    //  if(this.valid == "ok"){
    //     //Get Access
    //     await  this._accessGroups();
    //   }
    //   else if(this.valid == "Validok"){
    //Workflow Status Checking
    await this._checkWorkflowStatus();
    // }
  }
  //Get Parameter from URL
  private _queryParamGetting() {
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let documentindexid = params.get('did');

    if (documentindexid != "" && documentindexid != null) {
      this.documentIndexID = parseInt(documentindexid);
      this.valid = "ok";
    }
    else {
      this.setState({
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.invalidSendRequestLink, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(this.redirectUrl);
      }, 10000);
    }
  }
  //Get Access Groups
  private async _accessGroups() {
    let AccessGroup = [];
    let ok = "No";
    if (this.props.project) {
      AccessGroup = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.AccessGroups).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendApprovalWF'").get();
    }
    else {
      AccessGroup = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.AccessGroups).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendApprovalWF'").get();
    }

    let AccessGroupItems: any[] = AccessGroup[0].AccessGroups.split(',');
    console.log(AccessGroupItems);
    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).select("DepartmentID").get();
    console.log(DocumentIndexItem);
    let deptid = parseInt(DocumentIndexItem.DepartmentID);
    const DepartmentItem: any = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.DepartmentList).items.filter('Title eq ' + deptid).select("AccessGroups").get();
    let AG = DepartmentItem[0].AccessGroups;
    const AccessGroupItem: any = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.AccessGroupDetailsList).items.get();
    let AccessGroupID;
    console.log(AccessGroupItem.length);
    for (let a = 0; a < AccessGroupItem.length; a++) {
      if (AccessGroupItem[a].Title == AG) {
        AccessGroupID = AccessGroupItem[a].GroupID;
      }
    }
    //Logic App "Check Access Group"
    const postURL = "https://prod-05.uaecentral.logic.azure.com:443/workflows/60862323b80c44369d5bc091f5490bfa/triggers/manual/paths/invoke?api-version=2016-10-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=QzrPWl7wN5e6k873vy-X9qNeBk0VJojo1M7CzwslVsA";

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'Groupid': AccessGroupID,
      'CurrentUserMail': this.currentEmail

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
      if (responseJSON.ValidUser == "Yes") {
        this.valid = "Validok";
      }
      else {
        //  this.setState({
        //     access: "none",
        //     accessDeniedMsgBar: "",
        //     statusMessage: { isShowMessage: true, message: this.InvalidUser, messageType: 1 }
        //   });
        //   setTimeout(() => {
        //     this.setState({ accessDeniedMsgBar: 'none', });
        //     window.location.replace(this.RedirectUrl);
        //   }, 10000);
      }


    }
    else {
    }
  }
  //Workflow Status Checking
  private async _checkWorkflowStatus() {

    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).select("WorkflowStatus").get();
    if (DocumentIndexItem.WorkflowStatus == "Under Review" || DocumentIndexItem.WorkflowStatus == "Under Approval") {
      this.setState({
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.workflowStatus, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(this.redirectUrl);
      }, 10000);
    }
    else {
      this.setState({ access: "", accessDeniedMsgBar: "none", });
      await this._bindSendRequestForm();
    }
    if (this.props.project) {
      this.setState({ hideProject: false });
      await this._project();
    }
  }
  //Messages
  private async _userMessageSettings() {
    const userMessageSettings: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.userMessageSettings).items.select("Title,Message").filter("PageName eq 'SendRequest'").get();
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title == "InvalidSendRequestUser") {
        this.invalidUser = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title == "InvalidSendRequestLink") {
        this.invalidSendRequestLink = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title == "WorkflowStatusError") {
        this.workflowStatus = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title == "DccReview") {
        var DccReview = userMessageSettings[i].Message;
        this.DccReview = replaceString(DccReview, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title == "UnderApproval") {
        var UnderApproval = userMessageSettings[i].Message;
        this.UnderApproval = replaceString(UnderApproval, '[DocumentName]', this.state.documentName);

      }
      if (userMessageSettings[i].Title == "UnderReview") {
        var UnderReview = userMessageSettings[i].Message;
        this.UnderReview = replaceString(UnderReview, '[DocumentName]', this.state.documentName);

      }

    }

  }
  //Bind Send Request Form
  public async _bindSendRequestForm() {

    let DocumentID;
    let DocumentName;
    let OwnerName;
    let OwnerEmail;
    let OwnerId;
    let Revision;
    let LinkToDocument;
    let CriticalDocument;

    //Get Document Index
    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).select("DocumentID,DocumentName,Owner/ID,Owner/Title,Owner/EMail,Revision,SourceDocument,CriticalDocument,SourceDocumentID").expand("Owner").get();
    console.log(DocumentIndexItem);
    DocumentID = DocumentIndexItem.DocumentID;
    DocumentName = DocumentIndexItem.DocumentName;
    OwnerName = DocumentIndexItem.Owner.Title;
    OwnerId = DocumentIndexItem.Owner.ID;
    Revision = DocumentIndexItem.Revision;
    LinkToDocument = DocumentIndexItem.SourceDocument.Url;
    // this.SourceDocumentID = DocumentIndexItem.SourceDocumentID;
    CriticalDocument = DocumentIndexItem.CriticalDocument;
    this.setState({
      documentID: DocumentID,
      documentName: DocumentName,
      ownerName: OwnerName,
      ownerId: OwnerId,
      revision: Revision,
      linkToDoc: LinkToDocument,
      criticalDocument: CriticalDocument
    });
    const SourceDocumentItem: any = await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocumentLibrary).items.filter('DocumentIndexId eq ' + this.documentIndexID).get();
    console.log(SourceDocumentItem);
    this.sourceDocumentID = SourceDocumentItem[0].ID;
    await this._userMessageSettings();

  }
  public async _project() {
    let RevisionLevelArray = [];
    let sorted_RevisionLevel = [];
    let RevisionCoding;
    let TransmittalRevision;
    let AcceptanceCodeId;
    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).select("RevisionCodingId,TransmittalRevision,AcceptanceCodeId").get();
    console.log(DocumentIndexItem.RevisionCodingId);
    RevisionCoding = DocumentIndexItem.RevisionCodingId;
    AcceptanceCodeId = DocumentIndexItem.AcceptanceCodeId;
    TransmittalRevision = DocumentIndexItem.TransmittalRevision;
    const RevisionLevelItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.RevisionLevelList).items.select("Title,ID").get();
    console.log(RevisionLevelItem);
    for (let i = 0; i < RevisionLevelItem.length; i++) {

      let RevisionLevelItemdata = {
        key: RevisionLevelItem[i].ID,
        text: RevisionLevelItem[i].Title
      };
      RevisionLevelArray.push(RevisionLevelItemdata);

    }
    console.log(RevisionLevelArray);
    sorted_RevisionLevel = _.orderBy(RevisionLevelArray, 'text', ['asc']);
    this.setState({
      revisionLevelArray: sorted_RevisionLevel,
      revisionCoding: RevisionCoding,
      acceptanceCodeId: AcceptanceCodeId,
      transmittalRevision: TransmittalRevision
    });
    const projectInformation = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.projectInformationListName).items.get();
    console.log("projectInformation", projectInformation);
    if (projectInformation.length > 0) {
      for (var k in projectInformation) {
        if (projectInformation[k].Key == "ProjectName") {
          this.setState({
            projectName: projectInformation[k].Title,
          });
        }
        if (projectInformation[k].Key == "ProjectNumber") {
          this.setState({
            projectNumber: projectInformation[k].Title,
          });
        }
      }
    }
  }
  //Revision History Url
  private _openRevisionHistory = () => {
    window.open(this.props.siteUrl + "/SitePages/" + this.props.RevisionHistoryPage + ".aspx?ID=" + this.documentIndexID);
  }
  public async _revisionLevelChanged(option: { key: any; text: any }) {
    this.setState({ saveDisable: false });
    this.setState({ revisionLevelvalue: option.key });
  }
  public _reviewerChange = (items: any[]) => {
    this.setState({ saveDisable: false });
    console.log(items);
    this.getSelectedReviewers = [];
    for (let item in items) {
      this.getSelectedReviewers.push(items[item].id);
    }
    this.setState({ reviewers: this.getSelectedReviewers });
    console.log(this.getSelectedReviewers);


  }
  public _approverChange = (items: any[]) => {
    this.setState({ saveDisable: false });
    let approverEmail;
    let approverName;

    console.log(items);
    let getSelectedApprover = [];

    for (let item in items) {
      approverEmail = items[item].secondaryText,
        approverName = items[item].text,
        getSelectedApprover.push(items[item].id);
    }
    this.setState({
      approver: getSelectedApprover[0],
      approverEmail: approverEmail,
      approverName: approverName
    });
    console.log(approverEmail);


  }
  public _dccReviewerChange = (items: any[]) => {
    this.setState({ saveDisable: false });
    let dccreviewerEmail;
    let dccreviewerName;

    console.log(items);
    let getSelecteddccreviewer = [];

    for (let item in items) {
      dccreviewerEmail = items[item].secondaryText,
        dccreviewerName = items[item].text,
        getSelecteddccreviewer.push(items[item].id);
    }
    this.setState({
      dccReviewer: getSelecteddccreviewer[0],
      dccReviewerEmail: dccreviewerEmail,
      dccReviewerName: dccreviewerName
    });


  }
  private _onExpDatePickerChange = (date?: Date): void => {
    this.setState({ saveDisable: false });

    this.setState({ dueDate: date });

  }
  //Comment Change
  public _commentschange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ comments: comments, saveDisable: false });
  }
  private _submitSendRequest = async () => {
    this.setState({ saveDisable: true });
    let sorted_previousHeaderItems = [];
    let previousHeaderItem = 0;
    const previousHeaderItems = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.select("ID").filter("DocumentIndex eq '" + this.documentIndexID + "' and(WorkflowStatus eq 'Returned with comments')").get();
    if (previousHeaderItems.length != 0) {
      sorted_previousHeaderItems = _.orderBy(previousHeaderItems, 'ID', ['desc']);
      previousHeaderItem = sorted_previousHeaderItems[0].ID;
    }
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
      Title: this.state.documentID,
      Status: "Workflow Initiated",
      LogDate: this.today,
      WorkflowID: this.newheaderid,
      Revision: this.state.revision,
      DocumentIndexId: this.documentIndexID

    });
    if (this.props.project) {
      if (this.validator.fieldValid("Approver") && this.validator.fieldValid("DueDate") && this.validator.fieldValid("RevisionLevel")) {
        if (this.state.dccReviewer != "") {
          this._dccReview(previousHeaderItem);
        }
        else if (this.state.reviewers.length == 0) {
          this._underProjectApprove(previousHeaderItem);
        }
        else {
          this._underProjectReview(previousHeaderItem);
        }
        this.validator.hideMessages();
        this.setState({ requestSend: "" });
        setTimeout(() => this.setState({ requestSend: 'none' }), 3000);
        // window.location.replace(this.RedirectUrl);

        // this._onCancel();
      }

      else {
        this.validator.showMessages();
        this.forceUpdate();
      }

    }
    else {
      if (this.validator.fieldValid("Approver") && this.validator.fieldValid("DueDate")) {
        if (this.state.reviewers.length == 0) {
          this._underApprove(previousHeaderItem);
        }
        else {
          this._underReview(previousHeaderItem);

        }
        this.validator.hideMessages();
        this.setState({ requestSend: "" });
        setTimeout(() => this.setState({ requestSend: 'none' }), 3000);
        // window.location.replace(this.RedirectUrl);

        // this._onCancel();
      }

      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
  }
  public async _dccReview(previousHeaderItem) {
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: { results: this.state.reviewers },
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Review",
      DocumentControllerId: this.state.dccReviewer,
      RevisionLevelId: this.state.revisionLevelvalue,
      RevisionCodingId: this.state.revisionCoding,
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }).then(async i => {
      this.newheaderid = i.data.ID;
      //Task delegation getting user id from hubsite
      this.reqWeb.siteUsers.getByEmail(this.state.dccReviewerEmail).get().then(async user => {
        console.log('User Id: ', user.Id);
        this.setState({
          hubSiteUserId: user.Id,
        });
        //Task delegation 
        const taskDelegation: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "'").get();
        console.log(taskDelegation);
        if (taskDelegation.length > 0) {
          let duedate = moment(this.state.dueDate).toDate();
          let ToDate = moment(taskDelegation[0].ToDate).toDate();
          let FromDate = moment(taskDelegation[0].FromDate).toDate();
          duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
          ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
          FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
          if (moment(duedate).isBetween(FromDate, ToDate) || moment(duedate).isSame(FromDate) || moment(duedate).isSame(ToDate)) {
            this.setState({
              approverEmail: taskDelegation[0].DelegatedTo.EMail,
              approverName: taskDelegation[0].DelegatedTo.Title,

              delegatedToId: taskDelegation[0].DelegatedTo.ID,
              delegatedFromId: taskDelegation[0].DelegatedFor.ID,
            });
          }//duedate checking
        }
        //detail list adding an item for approval
        sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get().then(async DelegatedTo => {
          this.setState({
            delegateToIdInSubSite: DelegatedTo.Id,
          });
          sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get().then(async DelegatedFor => {
            this.setState({
              delegateForIdInSubSite: DelegatedFor.Id,
            });
            sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add
              ({
                HeaderIDId: Number(this.newheaderid),
                Workflow: "DCC Review",
                Title: this.state.documentName,
                ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.dccReviewer),
                DueDate: this.state.dueDate,
                DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                ResponseStatus: "Under Review"
              }).then(async r => {
                this.setState({ detailIdForApprover: r.data.ID });
                this.newDetailItemID = r.data.ID;
                sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
                  Link: {
                    "__metadata": { type: "SP.FieldUrlValue" },
                    Description: "Link to Review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + "&wf=dcc"
                  },
                });

                //MY tasks list updation
                await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.WorkflowTasksList).items.add
                  ({
                    Title: "Document Controller Review '" + this.state.documentName + "'",
                    Description: "DCC Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                    DueDate: this.state.dueDate,
                    StartDate: this.today,
                    AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : user.Id),
                    Workflow: "DCC Review",
                    // Priority:(this.state.criticalDocument == true ? "Critical" :""),
                    DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                    Source: (this.props.project ? "Project" : "QDMS"),
                    DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : 0),
                    Link: {
                      "__metadata": { type: "SP.FieldUrlValue" },
                      Description: "Link to Review",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + "&wf=dcc"
                    },

                  }).then(async taskId => {
                    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                      ({
                        TaskID: taskId.data.ID,
                      });

                    //notification preference checking                                 
                    this._sendmail(this.state.dccReviewerEmail, "DocDCCReview", this.state.dccReviewerName)
                      .then(aftermail => {
                        //Email pending  emailbody to approver                 

                        this.setState({
                          comments: "",
                          statusKey: "",
                          approverEmail: "",
                          approverName: "",
                          approver: "",
                        });
                      });//aftermail
                  });//taskID
              });//r
          });//DelegatedFor
        });//DelegatedTo
      }).then(async update => {
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
          WorkflowStatus: "Under Review",
          Workflow: "Review"
        });
        await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
          WorkflowStatus: "Under Review",
          Workflow: "Review"
        });
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
          Title: this.state.documentID,
          Status: "Under Review ",
          LogDate: this.today,
          WorkflowID: this.newheaderid,
          Revision: this.state.revision,
          Workflow: "DCC Review",
          DocumentIndexId: this.documentIndexID
        });
      }).then(msg => {

        setTimeout(() => {
          this.setState({
            statusMessage: { isShowMessage: true, message: this.DccReview, messageType: 4 },
          });
          window.location.replace(this.redirectUrl);
        }, 10000);
      });//msg
    });//newheaderid
  }
  public async _underApprove(previousHeaderItem) {

    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Approval",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      ReviewedDate: this.today,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: { results: this.state.currentUserRewviewer },
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Approve",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }).then(async i => {
      this.newheaderid = i.data.ID;
      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add({
        HeaderIDId: Number(this.newheaderid),
        Workflow: "Review",
        Title: this.state.documentName,
        ResponsibleId: this.currentId,
        DueDate: this.state.dueDate,
        ResponseStatus: "Reviewed",
        ResponseDate: this.today,
      }).then(async r => {
        this.setState({ detailIdForApprover: r.data.ID });
        this.newDetailItemID = r.data.ID;
        sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
          Link: {
            "__metadata": { type: "SP.FieldUrlValue" },
            Description: "Link to review",
            Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
          },
        });
      });
      //Task delegation getting user id from hubsite
      this.reqWeb.siteUsers.getByEmail(this.state.approverEmail).get().then(async user => {
        console.log('User Id: ', user.Id);
        this.setState({
          hubSiteUserId: user.Id,
        });

        //Task delegation 
        const taskDelegation: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "'").get();
        console.log(taskDelegation);
        if (taskDelegation.length > 0) {
          let duedate = moment(this.state.dueDate).toDate();
          let ToDate = moment(taskDelegation[0].ToDate).toDate();
          let FromDate = moment(taskDelegation[0].FromDate).toDate();
          duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
          ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
          FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
          if (moment(duedate).isBetween(FromDate, ToDate) || moment(duedate).isSame(FromDate) || moment(duedate).isSame(ToDate)) {
            this.setState({
              approverEmail: taskDelegation[0].DelegatedTo.EMail,
              approverName: taskDelegation[0].DelegatedTo.Title,

              delegatedToId: taskDelegation[0].DelegatedTo.ID,
              delegatedFromId: taskDelegation[0].DelegatedFor.ID,
            });
          }//duedate checking

          //detail list adding an item for approval
          sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get().then(async DelegatedTo => {
            this.setState({
              delegateToIdInSubSite: DelegatedTo.Id,
            });
            sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get().then(async DelegatedFor => {
              this.setState({
                delegateForIdInSubSite: DelegatedFor.Id,
              });
              sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add
                ({
                  HeaderIDId: Number(this.newheaderid),
                  Workflow: "Approval",
                  Title: this.state.documentName,
                  ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.approver),
                  DueDate: this.state.dueDate,
                  DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                  ResponseStatus: "Under Approval"
                }).then(async r => {
                  this.setState({ detailIdForApprover: r.data.ID });
                  this.newDetailItemID = r.data.ID;
                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
                    Link: {
                      "__metadata": { type: "SP.FieldUrlValue" },
                      Description: "Link to approve",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                    },
                  });

                  //MY tasks list updation
                  await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.WorkflowTasksList).items.add
                    ({
                      Title: "Approve '" + this.state.documentName + "'",
                      Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                      DueDate: this.state.dueDate,
                      StartDate: this.today,
                      AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : user.Id),
                      Workflow: "Approval",
                      // Priority:(this.state.criticalDocument == true ? "Critical" :""),
                      DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                      Source: (this.props.project ? "Project" : "QDMS"),
                      DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : 0),
                      Link: {
                        "__metadata": { type: "SP.FieldUrlValue" },
                        Description: "Link to approve",
                        Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                      },

                    }).then(taskId => {
                      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                        ({
                          TaskID: taskId.data.ID,
                        });
                      //notification preference checking                                 
                      this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName)
                        .then(aftermail => {
                          //Email pending  emailbody to approver                 

                          this.setState({
                            comments: "",
                            statusKey: "",
                            approverEmail: "",
                            approverName: "",
                            approver: "",
                          });
                        });//aftermail
                    });//taskID
                });//r

            });//DelegatedFor
          });//DelegatedTo
        }
        else {
          sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add
            ({
              HeaderIDId: Number(this.newheaderid),
              Workflow: "Approval",
              Title: this.state.documentName,
              ResponsibleId: this.state.approver,
              DueDate: this.state.dueDate,
              ResponseStatus: "Under Approval"
            }).then(async r => {
              this.setState({ detailIdForApprover: r.data.ID });
              this.newDetailItemID = r.data.ID;
              sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
                Link: {
                  "__metadata": { type: "SP.FieldUrlValue" },
                  Description: "Link to approve",
                  Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                },
              });

              //MY tasks list updation
              await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.WorkflowTasksList).items.add
                ({
                  Title: "Approve '" + this.state.documentName + "'",
                  Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                  DueDate: this.state.dueDate,
                  StartDate: this.today,
                  AssignedToId: user.Id,
                  Workflow: "Approval",
                  Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                  Source: (this.props.project ? "Project" : "QDMS"),
                  Link: {
                    "__metadata": { type: "SP.FieldUrlValue" },
                    Description: "Link to approve",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                  },

                }).then(taskId => {
                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                    ({
                      TaskID: taskId.data.ID,
                    });
                  //notification preference checking                                 
                  this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName)
                    .then(aftermail => {
                      //Email pending  emailbody to approver                 

                      this.setState({
                        comments: "",
                        statusKey: "",
                        approverEmail: "",
                        approverName: "",
                        approver: "",
                      });
                    });//aftermail
                });//taskID
            });//r
        }//else no delegation
      }).then(async update => {
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
          WorkflowStatus: "Under Approval",
          Workflow: "Approval"
        });
        await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
          WorkflowStatus: "Under Approval",
          Workflow: "Approval"
        });
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
          Title: this.state.documentID,
          Status: "Under Approval",
          LogDate: this.today,
          WorkflowID: this.newheaderid,
          Revision: this.state.revision,
          DocumentIndexId: this.documentIndexID,
          Workflow: "Approve"
        });
      }).then(msg => {

        setTimeout(() => {
          this.setState({
            statusMessage: { isShowMessage: true, message: this.UnderApproval, messageType: 4 },
          });
          window.location.replace(this.redirectUrl);
        }, 10000);
      });//msg
    });//newheaderid
  }
  public async _underReview(previousHeaderItem) {
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ReviewersId: { results: this.state.reviewers },
      ApproverId: this.state.approver,
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Review",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }).then(async i => {
      this.newheaderid = i.data.ID;
      //for reviewers if exist
      for (var k = 0; k <= this.state.reviewers.length; k++) {
        console.log(this.state.reviewers[k]);
        sp.web.siteUsers.getById(this.state.reviewers[k]).get().then(async user => {
          console.log(user);
          this.reqWeb.siteUsers.getByEmail(user.Email).get().then(async hubsieUser => {
            console.log(hubsieUser.Id);
            const taskDelegation: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + hubsieUser.Id + "'").get();
            console.log(taskDelegation);
            //Check if Task Delegation
            if (taskDelegation.length > 0) {
              let duedate = moment(this.state.dueDate).toDate();
              let ToDate = moment(taskDelegation[0].ToDate).toDate();
              let FromDate = moment(taskDelegation[0].FromDate).toDate();
              duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
              ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
              FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
              if (moment(duedate).isBetween(FromDate, ToDate) || moment(duedate).isSame(FromDate) || moment(duedate).isSame(ToDate)) {
                this.setState({
                  approverEmail: taskDelegation[0].DelegatedTo.EMail,
                  approverName: taskDelegation[0].DelegatedTo.Title,
                  delegatedToId: taskDelegation[0].DelegatedTo.ID,
                  delegatedFromId: taskDelegation[0].DelegatedFor.ID,
                });
              }
              //Get Delegated To ID
              sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get().then(async DelegatedTo => {

                this.setState({
                  delegateToIdInSubSite: DelegatedTo.Id,
                });
                //Get Delegated For ID
                sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get().then(async DelegatedFor => {

                  this.setState({
                    delegateForIdInSubSite: DelegatedFor.Id,
                  });
                  //detail list adding an item for reviewers

                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add({
                    HeaderIDId: Number(this.newheaderid),
                    Workflow: "Review",
                    Title: this.state.documentName,
                    ResponsibleId: (this.state.delegatedToId != "" ? DelegatedTo.Id : user.Id),
                    DueDate: this.state.dueDate,
                    DelegatedFromId: (this.state.delegatedToId != "" ? DelegatedFor.Id : parseInt("")),
                    ResponseStatus: "Under Review"
                  }).then(async r => {
                    this.setState({ detailIdForApprover: r.data.ID });
                    this.newDetailItemID = r.data.ID;
                    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
                      Link: {
                        "__metadata": { type: "SP.FieldUrlValue" },
                        Description: "Link to Review",
                        Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                      },
                    });//Update link
                    //MY tasks list updation with delegated from
                    await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.WorkflowTasksList).items.add({
                      Title: "Review '" + this.state.documentName + "'",
                      Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                      DueDate: this.state.dueDate,
                      StartDate: this.today,
                      AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : hubsieUser.Id),
                      Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                      DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                      Source: (this.props.project ? "Project" : "QDMS"),
                      DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : parseInt("")),
                      Workflow: "Review",
                      Link: {
                        "__metadata": { type: "SP.FieldUrlValue" },
                        Description: "Link to Review",
                        Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                      },
                    }).then(taskId => {
                      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                        ({
                          TaskID: taskId.data.ID,
                        });
                      this._sendmail(DelegatedTo.Email, "DocReview", DelegatedTo.Title);

                      this.setState({
                        statusMessage: { isShowMessage: true, message: this.UnderReview, messageType: 4 },
                        comments: "",
                        statusKey: "",
                        approverEmail: "",
                        approverName: "",
                        approver: "",
                        delegateForIdInSubSite: ""

                      });
                    });//taskID
                  });//r
                });//Delegated For
              });//Delegated To
            }//IF
            //If no task delegation
            else {
              //detail list adding an item for reviewers
              sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add({
                HeaderIDId: Number(this.newheaderid),
                Workflow: "Review",
                Title: this.state.documentName,
                ResponsibleId: user.Id,
                DueDate: this.state.dueDate,
                ResponseStatus: "Under Review"
              }).then(async r => {
                this.setState({ detailIdForApprover: r.data.ID });
                this.newDetailItemID = r.data.ID;
                sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
                  Link: {
                    "__metadata": { type: "SP.FieldUrlValue" },
                    Description: "Link to review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                  },
                });
                //MY tasks list updation with delegated from
                await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.WorkflowTasksList).items.add({
                  Title: "Review '" + this.state.documentName + "'",
                  Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                  DueDate: this.state.dueDate,
                  StartDate: this.today,
                  AssignedToId: hubsieUser.Id,
                  Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                  Source: (this.props.project ? "Project" : "QDMS"),
                  Workflow: "Review",
                  Link: {
                    "__metadata": { type: "SP.FieldUrlValue" },
                    Description: "Link to Review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                  },
                }).then(taskId => {
                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                    ({
                      TaskID: taskId.data.ID,
                    });
                  this._sendmail(user.Email, "DocReview", user.Title);

                  this.setState({
                    statusMessage: { isShowMessage: true, message: this.UnderReview, messageType: 4 },
                    comments: "",
                    statusKey: "",
                    approverEmail: "",
                    approverName: "",
                    approver: "",
                  });
                });//taskId
              });//r
            }//else
          });//hubsiteuser
        });//user
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
          WorkflowStatus: "Under Review",
          Workflow: "Review"
        });
        await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
          WorkflowStatus: "Under Review",
          Workflow: "Review"
        });
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
          Title: this.state.documentID,
          Status: "Under Review",
          LogDate: this.today,
          WorkflowID: this.newheaderid,
          Revision: this.state.revision,
          DocumentIndexId: this.documentIndexID,
          Workflow: "Review"
        }).then(msg => {

          setTimeout(() => {
            this.setState({
              statusMessage: { isShowMessage: true, message: this.UnderReview, messageType: 4 },
            });
            window.location.replace(this.redirectUrl);
          }, 10000);
        });//msg
      }
    });
  }
  public async _underProjectApprove(previousHeaderItem) {
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Approval",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      ReviewedDate: this.today,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: { results: this.state.currentUserRewviewer },
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Approve",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString(),
      RevisionLevelId: this.state.revisionLevelvalue,
      RevisionCodingId: this.state.revisionCoding,
      TransmittalRevision: this.state.transmittalRevision,
      AcceptanceCodeId: this.state.acceptanceCodeId

    }).then(async i => {
      this.newheaderid = i.data.ID;
      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add({
        HeaderIDId: Number(this.newheaderid),
        Workflow: "Review",
        Title: this.state.documentName,
        ResponsibleId: this.currentId,
        DueDate: this.state.dueDate,
        ResponseStatus: "Reviewed",
        ResponseDate: this.today,
      }).then(async r => {
        this.setState({ detailIdForApprover: r.data.ID });
        this.newDetailItemID = r.data.ID;
        sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
          Link: {
            "__metadata": { type: "SP.FieldUrlValue" },
            Description: "Link to review",
            Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
          },
        });
      });
      //Task delegation getting user id from hubsite
      this.reqWeb.siteUsers.getByEmail(this.state.approverEmail).get().then(async user => {
        console.log('User Id: ', user.Id);
        this.setState({
          hubSiteUserId: user.Id,
        });

        //Task delegation 
        const taskDelegation: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "'").get();
        console.log(taskDelegation);
        if (taskDelegation.length > 0) {
          let duedate = moment(this.state.dueDate).toDate();
          let ToDate = moment(taskDelegation[0].ToDate).toDate();
          let FromDate = moment(taskDelegation[0].FromDate).toDate();
          duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
          ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
          FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
          if (moment(duedate).isBetween(FromDate, ToDate) || moment(duedate).isSame(FromDate) || moment(duedate).isSame(ToDate)) {
            this.setState({
              approverEmail: taskDelegation[0].DelegatedTo.EMail,
              approverName: taskDelegation[0].DelegatedTo.Title,

              delegatedToId: taskDelegation[0].DelegatedTo.ID,
              delegatedFromId: taskDelegation[0].DelegatedFor.ID,
            });
          }//duedate checking

          //detail list adding an item for approval
          sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get().then(async DelegatedTo => {
            this.setState({
              delegateToIdInSubSite: DelegatedTo.Id,
            });
            sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get().then(async DelegatedFor => {
              this.setState({
                delegateForIdInSubSite: DelegatedFor.Id,
              });
              sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add
                ({
                  HeaderIDId: Number(this.newheaderid),
                  Workflow: "Approval",
                  Title: this.state.documentName,
                  ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.approver),
                  DueDate: this.state.dueDate,
                  DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegateForIdInSubSite : parseInt("")),
                  ResponseStatus: "Under Approval"
                }).then(async r => {
                  this.setState({ detailIdForApprover: r.data.ID });
                  this.newDetailItemID = r.data.ID;
                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
                    Link: {
                      "__metadata": { type: "SP.FieldUrlValue" },
                      Description: "Link to approve",
                      Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                    },
                  });

                  //MY tasks list updation
                  await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.WorkflowTasksList).items.add
                    ({
                      Title: "Approve '" + this.state.documentName + "'",
                      Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                      DueDate: this.state.dueDate,
                      StartDate: this.today,
                      AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : user.Id),
                      Workflow: "Approval",
                      // Priority:(this.state.criticalDocument == true ? "Critical" :""),
                      DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                      Source: (this.props.project ? "Project" : "QDMS"),
                      DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : 0),
                      Link: {
                        "__metadata": { type: "SP.FieldUrlValue" },
                        Description: "Link to approve",
                        Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                      },

                    }).then(taskId => {
                      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                        ({
                          TaskID: taskId.data.ID,
                        });
                      //notification preference checking                                 
                      this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName)
                        .then(aftermail => {
                          //Email pending  emailbody to approver                 

                          this.setState({
                            comments: "",
                            statusKey: "",
                            approverEmail: "",
                            approverName: "",
                            approver: "",
                          });
                        });//aftermail
                    });//taskID
                });//r

            });//DelegatedFor
          });//DelegatedTo
        }
        else {
          sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add
            ({
              HeaderIDId: Number(this.newheaderid),
              Workflow: "Approval",
              Title: this.state.documentName,
              ResponsibleId: this.state.approver,
              DueDate: this.state.dueDate,
              ResponseStatus: "Under Approval"
            }).then(async r => {
              this.setState({ detailIdForApprover: r.data.ID });
              this.newDetailItemID = r.data.ID;
              sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
                Link: {
                  "__metadata": { type: "SP.FieldUrlValue" },
                  Description: "Link to approve",
                  Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                },
              });

              //MY tasks list updation
              await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.WorkflowTasksList).items.add
                ({
                  Title: "Approve '" + this.state.documentName + "'",
                  Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                  DueDate: this.state.dueDate,
                  StartDate: this.today,
                  AssignedToId: user.Id,
                  Workflow: "Approval",
                  Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                  Source: (this.props.project ? "Project" : "QDMS"),
                  Link: {
                    "__metadata": { type: "SP.FieldUrlValue" },
                    Description: "Link to approve",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentApprovalPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                  },

                }).then(taskId => {
                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                    ({
                      TaskID: taskId.data.ID,
                    });
                  //notification preference checking                                 
                  this._sendmail(this.state.approverEmail, "DocApproval", this.state.approverName)
                    .then(aftermail => {
                      //Email pending  emailbody to approver                 

                      this.setState({
                        comments: "",
                        statusKey: "",
                        approverEmail: "",
                        approverName: "",
                        approver: "",
                      });
                    });//aftermail
                });//taskID
            });//r
        }//else no delegation
      }).then(async update => {
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
          WorkflowStatus: "Under Approval",
          Workflow: "Approval"
        });
        await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
          WorkflowStatus: "Under Approval",
          Workflow: "Approval"
        });
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
          Title: this.state.documentID,
          Status: "Under Approval",
          LogDate: this.today,
          WorkflowID: this.newheaderid,
          Revision: this.state.revision,
          DocumentIndexId: this.documentIndexID,
          Workflow: "Approve"
        });
      }).then(msg => {

        setTimeout(() => {
          this.setState({
            statusMessage: { isShowMessage: true, message: this.UnderApproval, messageType: 4 },
          });
          window.location.replace(this.redirectUrl);
        }, 10000);
      });//msg
    });//newheaderid
  }
  public async _underProjectReview(previousHeaderItem) {
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.sourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ReviewersId: { results: this.state.reviewers },
      ApproverId: this.state.approver,
      RequesterId: this.currentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Review",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString(),
      RevisionLevelId: this.state.revisionLevelvalue,
      RevisionCodingId: this.state.revisionCoding,
      TransmittalRevision: this.state.transmittalRevision,
      AcceptanceCodeId: this.state.acceptanceCodeId

    }).then(async i => {
      this.newheaderid = i.data.ID;
      //for reviewers if exist
      for (var k = 0; k <= this.state.reviewers.length; k++) {
        console.log(this.state.reviewers[k]);
        sp.web.siteUsers.getById(this.state.reviewers[k]).get().then(async user => {
          console.log(user);
          this.reqWeb.siteUsers.getByEmail(user.Email).get().then(async hubsieUser => {
            console.log(hubsieUser.Id);
            const taskDelegation: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + hubsieUser.Id + "'").get();
            console.log(taskDelegation);
            //Check if Task Delegation
            if (taskDelegation.length > 0) {
              let duedate = moment(this.state.dueDate).toDate();
              let ToDate = moment(taskDelegation[0].ToDate).toDate();
              let FromDate = moment(taskDelegation[0].FromDate).toDate();
              duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
              ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
              FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
              if (moment(duedate).isBetween(FromDate, ToDate) || moment(duedate).isSame(FromDate) || moment(duedate).isSame(ToDate)) {
                this.setState({
                  approverEmail: taskDelegation[0].DelegatedTo.EMail,
                  approverName: taskDelegation[0].DelegatedTo.Title,
                  delegatedToId: taskDelegation[0].DelegatedTo.ID,
                  delegatedFromId: taskDelegation[0].DelegatedFor.ID,
                });
              }
              //Get Delegated To ID
              sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedTo.EMail).get().then(async DelegatedTo => {

                this.setState({
                  delegateToIdInSubSite: DelegatedTo.Id,
                });
                //Get Delegated For ID
                sp.web.siteUsers.getByEmail(taskDelegation[0].DelegatedFor.EMail).get().then(async DelegatedFor => {

                  this.setState({
                    delegateForIdInSubSite: DelegatedFor.Id,
                  });
                  //detail list adding an item for reviewers

                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add({
                    HeaderIDId: Number(this.newheaderid),
                    Workflow: "Review",
                    Title: this.state.documentName,
                    ResponsibleId: (this.state.delegatedToId != "" ? DelegatedTo.Id : user.Id),
                    DueDate: this.state.dueDate,
                    DelegatedFromId: (this.state.delegatedToId != "" ? DelegatedFor.Id : parseInt("")),
                    ResponseStatus: "Under Review"
                  }).then(async r => {
                    this.setState({ detailIdForApprover: r.data.ID });
                    this.newDetailItemID = r.data.ID;
                    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
                      Link: {
                        "__metadata": { type: "SP.FieldUrlValue" },
                        Description: "Link to Review",
                        Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                      },
                    });//Update link
                    //MY tasks list updation with delegated from
                    await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.WorkflowTasksList).items.add({
                      Title: "Review '" + this.state.documentName + "'",
                      Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                      DueDate: this.state.dueDate,
                      StartDate: this.today,
                      AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : hubsieUser.Id),
                      Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                      DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                      Source: (this.props.project ? "Project" : "QDMS"),
                      DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : parseInt("")),
                      Workflow: "Review",
                      Link: {
                        "__metadata": { type: "SP.FieldUrlValue" },
                        Description: "Link to Review",
                        Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                      },
                    }).then(taskId => {
                      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                        ({
                          TaskID: taskId.data.ID,
                        });
                      this._sendmail(DelegatedTo.Email, "DocReview", DelegatedTo.Title);

                      this.setState({
                        statusMessage: { isShowMessage: true, message: this.UnderReview, messageType: 4 },
                        comments: "",
                        statusKey: "",
                        approverEmail: "",
                        approverName: "",
                        approver: "",
                        delegateForIdInSubSite: ""

                      });
                    });//taskID
                  });//r
                });//Delegated For
              });//Delegated To
            }//IF
            //If no task delegation
            else {
              //detail list adding an item for reviewers
              sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.add({
                HeaderIDId: Number(this.newheaderid),
                Workflow: "Review",
                Title: this.state.documentName,
                ResponsibleId: user.Id,
                DueDate: this.state.dueDate,
                ResponseStatus: "Under Review"
              }).then(async r => {
                this.setState({ detailIdForApprover: r.data.ID });
                this.newDetailItemID = r.data.ID;
                sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update({
                  Link: {
                    "__metadata": { type: "SP.FieldUrlValue" },
                    Description: "Link to review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                  },
                });
                //MY tasks list updation with delegated from
                await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.WorkflowTasksList).items.add({
                  Title: "Review '" + this.state.documentName + "'",
                  Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentUser + "' on '" + this.today + "'",
                  DueDate: this.state.dueDate,
                  StartDate: this.today,
                  AssignedToId: hubsieUser.Id,
                  Priority: (this.state.criticalDocument == true ? "Critical" : ""),
                  Source: (this.props.project ? "Project" : "QDMS"),
                  Workflow: "Review",
                  Link: {
                    "__metadata": { type: "SP.FieldUrlValue" },
                    Description: "Link to Review",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.DocumentReviewPage + ".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                  },
                }).then(taskId => {
                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                    ({
                      TaskID: taskId.data.ID,
                    });
                  this._sendmail(user.Email, "DocReview", user.Title);

                  this.setState({
                    statusMessage: { isShowMessage: true, message: this.UnderReview, messageType: 4 },
                    comments: "",
                    statusKey: "",
                    approverEmail: "",
                    approverName: "",
                    approver: "",
                  });
                });//taskId
              });//r
            }//else
          });//hubsiteuser
        });//user
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
          WorkflowStatus: "Under Review",
          Workflow: "Review"
        });
        await sp.web.getList(this.props.siteUrl + "/" + this.props.SourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
          WorkflowStatus: "Under Review",
          Workflow: "Review"
        });
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
          Title: this.state.documentID,
          Status: "Under Review",
          LogDate: this.today,
          WorkflowID: this.newheaderid,
          Revision: this.state.revision,
          DocumentIndexId: this.documentIndexID,
          Workflow: "Review"
        }).then(msg => {

          setTimeout(() => {
            this.setState({
              statusMessage: { isShowMessage: true, message: this.UnderReview, messageType: 4 },
            });
            window.location.replace(this.redirectUrl);
          }, 10000);
        });//msg
      }
    });
  }
  //Send Mail
  public _sendmail = async (emailuser, type, name) => {

    let formatday = moment(this.today).format('DD/MMM/YYYY');
    let day = formatday.toString();
    let mailSend = "No";
    let Subject;
    let Body;
    console.log(this.state.criticalDocument);

    const notificationPreference: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.notificationPreference).items.filter("EmailUser/EMail eq '" + emailuser + "'").select("Preference").get();
    console.log(notificationPreference[0].Preference);
    if (notificationPreference.length > 0) {
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
    else if (this.state.criticalDocument == true) {
      //console.log("Send mail for critical document");
      mailSend = "Yes";
    }

    if (mailSend == "Yes") {
      const emailNotification: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.emailNotification).items.get();
      console.log(emailNotification);
      for (var k in emailNotification) {
        if (emailNotification[k].Title == type) {
          Subject = emailNotification[k].Subject;
          Body = emailNotification[k].Body;
        }

      }
      let replacedSubject = replaceString(Subject, '[DocumentName]', this.state.documentName);
      let replaceRequester = replaceString(Body, '[Sir/Madam]', name);
      let replaceBody = replaceString(replaceRequester, '[DocumentName]', this.state.documentName);

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
  public render(): React.ReactElement<IEmecSendRequestProps> {
    return (
      <div className={styles.emecSendRequest}>
        <Desktop>
          <div style={{ display: this.state.access }}>
            <div className={styles.alignCenter}> Review and approval request form</div>
            <br></br>
            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>
            <div className={styles.flex}>
              <div className={styles.width}><Label >Document ID : {this.state.documentID}</Label></div>
              <div ><Link onClick={this._openRevisionHistory} underline>Revision History</Link></div>
            </div>
            <div hidden={this.state.hideProject}>
              <div className={styles.flex}>
                <div className={styles.width}><Label >Project Name : {this.state.projectName} </Label></div>
                <div><Label >Project Number : {this.state.projectNumber}</Label></div>
              </div>
            </div>
            <div className={styles.flex}>
              <div className={styles.width}><Label >Document : <a href={this.state.linkToDoc}>{this.state.documentName}</a></Label></div>
              <div ><Label >Revision : {this.state.revision}</Label></div>
            </div>
            <div className={styles.flex}>
              <div className={styles.width}><Label >Owner : {this.state.ownerName} </Label></div>
              <div><Label >Requester : {this.state.currentUser}</Label></div>
            </div>
            <div hidden={this.state.hideProject}>

              <div className={styles.flex} >
                <div className={styles.width} style={{ paddingRight: '15px' }}>
                  <Dropdown
                    placeholder="Select Option"
                    label="RevisionLevel"
                    style={{ marginBottom: '10px', backgroundColor: "white", height: '34px' }}
                    options={this.state.revisionLevelArray}
                    onChanged={this._revisionLevelChanged}
                    selectedKey={this.state.revisionLevelvalue}
                    required />
                  <div style={{ color: "#dc3545" }}>{this.validator.message("RevisionLevel", this.state.revisionLevelvalue, "required")}{" "}</div>
                </div>
                <div className={styles.width}>
                  <PeoplePicker
                    context={this.props.context}
                    titleText="DCC"
                    personSelectionLimit={1}
                    groupName={""} // Leave this blank in case you want to filter from all users    
                    showtooltip={true}
                    disabled={false}
                    ensureUser={true}

                    selectedItems={(items) => this._dccReviewerChange(items)}
                    defaultSelectedUsers={[this.state.dcc]}
                    showHiddenInUI={false}
                    // isRequired={true}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                  /></div>
              </div>
            </div>
            <div><PeoplePicker
              context={this.props.context}
              titleText="Reviewer(s)"
              personSelectionLimit={20}
              groupName={""} // Leave this blank in case you want to filter from all users    
              showtooltip={true}
              disabled={false}
              ensureUser={true}
              selectedItems={(items) => this._reviewerChange(items)}
              defaultSelectedUsers={[this.state.reviewer]}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            /></div>

            <div className={styles.flex}>
              <div className={styles.width}>
                <PeoplePicker
                  context={this.props.context}
                  titleText="Approver"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users    
                  showtooltip={true}
                  disabled={false}
                  ensureUser={true}
                  selectedItems={(items) => this._approverChange(items)}
                  //defaultSelectedUsers={[this.state.approver]}
                  showHiddenInUI={false}
                  isRequired={true}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                /> <div style={{ color: "#dc3545" }}>{this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
              </div>
              <div className={styles.width} style={{ paddingLeft: '15px', marginTop: '2px' }}>
                <DatePicker label="DueDate:" id="DueDate"
                  onSelectDate={this._onExpDatePickerChange}
                  placeholder="Select a date..."
                  isRequired={true}
                  value={this.state.dueDate}
                  minDate={new Date()}
                // className={controlClass.control}
                // onSelectDate={this._onDatePickerChange}                 
                /><div style={{ color: "#dc3545" }}>{this.validator.message("DueDate", this.state.dueDate, "required")}{" "}</div>
              </div>
            </div>

            <div className={styles.mt}>
              < TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentschange} multiline required autoAdjustHeight></TextField></div>
            <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div>

            <DialogFooter>

              <div className={styles.rgtalign}>
                <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
              </div>
              <div className={styles.rgtalign} >
                <DefaultButton id="b2" className={styles.btn} disabled={this.state.saveDisable} onClick={this._submitSendRequest}>Submit</DefaultButton >
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
            <div className={styles.alignCenter}> Review and approval request form</div>
            <br></br>
            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>

            <div className={styles.width}><Label >Document ID : {this.state.documentID}</Label></div>
            <div ><Link onClick={this._openRevisionHistory} underline>Revision History</Link></div>
            <div hidden={this.state.hideProject}>
              <div className={styles.width}><Label >Project Name : {this.state.projectName} </Label></div>
              <div><Label >Project Number : {this.state.projectNumber}</Label></div>
            </div>

            <div className={styles.width}><Label >Document : <a href={this.state.linkToDoc}>{this.state.documentName}</a></Label></div>
            <div ><Label >Revision : {this.state.revision}</Label></div>


            <div className={styles.width}><Label >Owner : {this.state.ownerName} </Label></div>
            <div><Label >Requester : {this.state.currentUser}</Label></div>

            <div hidden={this.state.hideProject}>

              <div className={styles.drpdwn} >
                <Dropdown
                  placeholder="Select Option"
                  label="RevisionLevel"
                  style={{ marginBottom: '10px', backgroundColor: "white" }}
                  options={this.state.revisionLevelArray}
                  onChanged={this._revisionLevelChanged}
                  selectedKey={this.state.revisionLevelvalue}
                  required />
                <div style={{ color: "#dc3545" }}>{this.validator.message("RevisionLevel", this.state.revisionLevelvalue, "required")}{" "}</div>
              </div>

              <PeoplePicker
                context={this.props.context}
                titleText="DCC"
                personSelectionLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users    
                showtooltip={true}
                disabled={false}
                ensureUser={true}
                selectedItems={(items) => this._dccReviewerChange(items)}
                defaultSelectedUsers={[this.state.dcc]}
                showHiddenInUI={false}
                // isRequired={true}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />

            </div>
            <div><PeoplePicker
              context={this.props.context}
              titleText="Reviewer(s)"
              personSelectionLimit={20}
              groupName={""} // Leave this blank in case you want to filter from all users    
              showtooltip={true}
              disabled={false}
              ensureUser={true}
              selectedItems={(items) => this._reviewerChange(items)}
              defaultSelectedUsers={[this.state.reviewer]}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            /></div>


            <div className={styles.drpdwn}>
              <PeoplePicker
                context={this.props.context}
                titleText="Approver"
                personSelectionLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users    
                showtooltip={true}
                disabled={false}
                ensureUser={true}
                selectedItems={(items) => this._approverChange(items)}
                //defaultSelectedUsers={[this.state.approver]}
                showHiddenInUI={false}
                isRequired={true}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              /> <div style={{ color: "#dc3545" }}>{this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
            </div>
            <div>
              <DatePicker label="DueDate:" id="DueDate"
                onSelectDate={this._onExpDatePickerChange}
                placeholder="Select a date..."
                isRequired={true}
                value={this.state.dueDate}
                minDate={new Date()}
              // className={controlClass.control}
              // onSelectDate={this._onDatePickerChange}                 
              /><div style={{ color: "#dc3545" }}>{this.validator.message("DueDate", this.state.dueDate, "required")}{" "}</div>
            </div>

            <div className={styles.mt}>
              < TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentschange} multiline required autoAdjustHeight></TextField></div>
            <div style={{ color: "#dc3545" }}>{this.validator.message("comments", this.state.comments, "required")}{" "}</div>
            <DialogFooter>

              <div className={styles.rgtalign}>
                <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
              </div>
              <div className={styles.rgtalign} >
                <DefaultButton id="b2" className={styles.btn} disabled={this.state.saveDisable} onClick={this._submitSendRequest}>Submit</DefaultButton >
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
