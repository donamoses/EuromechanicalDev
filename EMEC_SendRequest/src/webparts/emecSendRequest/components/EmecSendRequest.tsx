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
  currentuser: string;
  hideproject: boolean;
  RevisionLevel: any[];
  RevisionLevelvalue: any;
  dcc: any;
  reviewer: any;
  dueDate: any;
  approver: any;
  comments: any;
  cancelConfirmMsg: string;
  confirmDialog: boolean;
  savedisable: boolean;
  RequestSend: string;
  statusKey: string;
  access: any;
  accessDeniedMsgBar: any;
  reviewers: any[];
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
  CriticalDocument: any;
  dccreviewerName:any;
  dccreviewerEmail:any;
  dccreviewer:any;
  RevisionLevelArray:any[];
  RevisionCoding:any;

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
  private InvalidUser;
  private CurrentEmail;
  private CurrentId;
  private today;
  private time;
  private WorkflowStatus;
  private SourceDocumentID;
  private newheaderid;
  private newDetailItemID;
  private DccReview;
  private UnderApproval;
  private UnderReview;
  private RedirectUrl = this.props.siteUrl + this.props.RedirectUrl;
  private InvalidSendRequestLink;
  
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
      currentuser: "",
      hideproject: true,
      RevisionLevel: [],
      RevisionLevelvalue: "",
      dcc: "",
      reviewer: "",
      dueDate: "",
      approver: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
      savedisable: false,
      RequestSend: 'none',
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
      CriticalDocument: "",
      dccreviewerName:"",
      dccreviewerEmail:"",
      dccreviewer:"",
      RevisionLevelArray:[],
      RevisionCoding:""

    };
    this.componentDidMount = this.componentDidMount.bind(this);
    this.userMessageSettings = this.userMessageSettings.bind(this);
    this.queryParamGetting = this.queryParamGetting.bind(this);
    this._openRevisionHistory = this._openRevisionHistory.bind(this);
    this.BindSendRequestForm = this.BindSendRequestForm.bind(this);
    this.project = this.project.bind(this);
    this.RevisionLevelChanged = this.RevisionLevelChanged.bind(this);
    this._dccReviewerChange = this._dccReviewerChange.bind(this);
    this._ReviewerChange = this._ReviewerChange.bind(this);
    this._ApproverChange = this._ApproverChange.bind(this);
    this._submitSendRequest = this._submitSendRequest.bind(this);
    this.dccReview = this.dccReview.bind(this);
    this.underApprove = this.underApprove.bind(this);
    this.underReview = this.underReview.bind(this);
    this.underProjectApprove = this.underProjectApprove.bind(this);
    this.underProjectReview = this.underProjectReview.bind(this);
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
    await this.userMessageSettings();
    //Get Parameter from URL
    this.queryParamGetting();

    if (this.props.project) {
      this.setState({ hideproject: false });
    }
    //Get Current User
    const user = await sp.web.currentUser.get();
    this.CurrentEmail = user.Email;
    this.CurrentId = user.Id;
    //Get Today
    this.today = new Date();
    this.setState({ currentuser: user.Title });
    //Get Access
    // this.accessGroups();
   
    //Workflow Status Checking
    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).select("WorkflowStatus").get();
    if (DocumentIndexItem.WorkflowStatus == "Under Review" || DocumentIndexItem.WorkflowStatus == "Under Approval") {
      this.setState({
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.WorkflowStatus, messageType: 1 },
      });
      setTimeout(() => {
        this.setState({ accessDeniedMsgBar: 'none', });
        window.location.replace(this.RedirectUrl);
      }, 5000);
    }
    else {
      this.setState({ access: "", accessDeniedMsgBar: "none", });
      await this.BindSendRequestForm();
    }
    if (this.props.project) {
      this.setState({ hideproject: false });
      await this.project();
    }

  }
  //Get Parameter from URL
  private queryParamGetting() {
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let documentindexid = params.get('did');

    if (documentindexid != "" && documentindexid != null) {
      this.documentIndexID = parseInt(documentindexid);

    }
    else {
      this.setState({
        accessDeniedMsgBar: "",
        statusMessage: { isShowMessage: true, message: this.InvalidSendRequestLink, messageType: 1 },
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
      AccessGroup= await this.reqWeb.lists.getByTitle(this.props.AccessGroups).items.select("AccessGroups,AccessFields").filter("Title eq 'QDMS_SendApprovalWF'").get();
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
  //User in HO Group
  try {
    let grp1: any[] = await sp.web.siteGroups.getByName(AccessGroupItems[result]).users();
    for (let i = 0; i < grp1.length; i++) {
        if (this.CurrentEmail == grp1[i].Email) {
            ok = "Yes";
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
  //Messages
  private async userMessageSettings() {
    const userMessageSettings: any[] = await this.reqWeb.lists.getByTitle(this.props.userMessageSettings).items.select("Title,Message").filter("PageName eq 'SendRequest'").get();
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title == "InvalidSendRequestUser") {
        this.InvalidUser = userMessageSettings[i].Message;
      }
      else if (userMessageSettings[i].Title == "InvalidSendRequestLink") {
        this.InvalidSendRequestLink = userMessageSettings[i].Message;
      }
      if (userMessageSettings[i].Title == "WorkflowStatusError") {
        this.WorkflowStatus = userMessageSettings[i].Message;
      }
      if(userMessageSettings[i].Title == "DccReview"){
        var DccReview = userMessageSettings[i].Message;
        this.DccReview = replaceString(DccReview, '[DocumentName]', this.state.documentName);
        
      }
      if(userMessageSettings[i].Title == "UnderApproval"){
        var UnderApproval = userMessageSettings[i].Message;
        this.UnderApproval = replaceString(UnderApproval, '[DocumentName]', this.state.documentName);
        
      }
      if(userMessageSettings[i].Title == "UnderReview"){
        var UnderReview = userMessageSettings[i].Message;
        this.UnderReview = replaceString(UnderReview, '[DocumentName]', this.state.documentName);
        
      }

    }

  }
//Bind Send Request Form
  public async BindSendRequestForm() {

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
    this.SourceDocumentID = DocumentIndexItem.SourceDocumentID;
    CriticalDocument = DocumentIndexItem.CriticalDocument;
    this.setState({
      documentID: DocumentID,
      documentName: DocumentName,
      ownerName: OwnerName,
      ownerId: OwnerId,
      revision: Revision,
      linkToDoc: LinkToDocument,
      CriticalDocument: CriticalDocument
    });
    await this.userMessageSettings();
  }
  public async project(){
    let RevisionLevelArray = [];
    let sorted_RevisionLevel = [];
    let RevisionCoding;
    const DocumentIndexItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentIndexList).items.getById(this.documentIndexID).select("RevisionCodingId").get();
    console.log(DocumentIndexItem.RevisionCodingId);
    RevisionCoding = DocumentIndexItem.RevisionCodingId;
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
               RevisionLevelArray:sorted_RevisionLevel,
               RevisionCoding:RevisionCoding
            });
             
  }
  //Revision History Url
  private _openRevisionHistory = () => {
    window.open(this.props.siteUrl + "/SitePages/"+this.props.RevisionHistoryPage+".aspx?ID=" + this.documentIndexID);
  }
  public async RevisionLevelChanged(option: { key: any; text: any }) {
    this.setState({ RevisionLevelvalue: option.key });
  }
  public _ReviewerChange = (items: any[]) => {
    
    console.log(items);
    let getSelectedReviewers = [];

    for (let item in items) {
     
      getSelectedReviewers.push(items[item].id);
    }
    this.setState({ reviewers: getSelectedReviewers });
    console.log(getSelectedReviewers);


  }
  public _ApproverChange = (items: any[]) => {
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
  public _dccReviewerChange =(items: any[]) =>{
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
      dccreviewer: getSelecteddccreviewer[0],
      dccreviewerEmail: dccreviewerEmail,
      dccreviewerName: dccreviewerName
    });
    

  }
  private _onExpDatePickerChange = (date?: Date): void => {

       
    this.setState({ dueDate: date });

  }
  //Comment Change
  public _commentschange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ comments: comments, savedisable: false });
  }
  private _submitSendRequest = async () => {
    let sorted_previousHeaderItems = [];
    let previousHeaderItem = 0;
    const previousHeaderItems = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.select("ID").filter("DocumentIndex eq '" + this.documentIndexID + "' and(WorkflowStatus eq 'Returned with comments')").get();
    if (previousHeaderItems.length != 0) {
      sorted_previousHeaderItems = _.orderBy(previousHeaderItems, 'ID', ['desc']);
      previousHeaderItem = sorted_previousHeaderItems[0].ID;
    }
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
      Title:this.state.documentID,
      Status:"Workflow Initiated",
      LogDate:this.today,
      WorkflowID:this.newheaderid,
      Revision:this.state.revision,
      DocumentIndexId:this.documentIndexID
     
    });
    if(this.props.project){
      if (this.validator.fieldValid("Approver") && this.validator.fieldValid("DueDate")&& this.validator.fieldValid("RevisionLevel")) {
        if(this.state.dccreviewer!=""){
          this.dccReview(previousHeaderItem);
        }
        else if(this.state.reviewers.length == 0) {
          this.underProjectApprove(previousHeaderItem);
        }
        else {
          this.underProjectReview(previousHeaderItem);
        }
        this.validator.hideMessages();
        this.setState({ RequestSend: "" });
        setTimeout(() => this.setState({ RequestSend: 'none' }), 3000);
        // window.location.replace(this.RedirectUrl);
  
        // this._onCancel();
      }
  
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    
    }
else{
    if (this.validator.fieldValid("Approver") && this.validator.fieldValid("DueDate")) {
      if (this.state.reviewers.length == 0) {
        this.underApprove(previousHeaderItem);
      }
      else {
        this.underReview(previousHeaderItem);

      }
      this.validator.hideMessages();
      this.setState({ RequestSend: "" });
      setTimeout(() => this.setState({ RequestSend: 'none' }), 3000);
      // window.location.replace(this.RedirectUrl);

      // this._onCancel();
    }

    else {
      this.validator.showMessages();
      this.forceUpdate();
    }
  }
  }
  public async dccReview(previousHeaderItem){
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.SourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: { results: this.state.reviewers },
      RequesterId: this.CurrentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Review",
      DocumentControllerId:this.state.dccreviewer,
      RevisionLevelId:this.state.RevisionLevelvalue,
      RevisionCodingId:this.state.RevisionCoding,
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }).then(async i => {
      this.newheaderid = i.data.ID;
       //Task delegation getting user id from hubsite
    this.reqWeb.siteUsers.getByEmail(this.state.dccreviewerEmail).get().then(async user => {
      console.log('User Id: ', user.Id);
      this.setState({
        hubSiteUserId: user.Id,
      });
      //Task delegation 
      const taskDelegation: any[] = await this.reqWeb.lists.getByTitle(this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "'").get();
      console.log(taskDelegation);
      if (taskDelegation.length > 0) {
        let duedate = moment(this.state.dueDate).toDate();
        let ToDate = moment(taskDelegation[0].ToDate).toDate();
        let FromDate = moment(taskDelegation[0].FromDate).toDate();
        duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
        ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
        FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
        if (moment(duedate).isBetween(FromDate, ToDate) ||moment(duedate).isSame(FromDate)||moment(duedate).isSame(ToDate)) {
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
              ResponsibleId: (this.state.delegatedToId != "" ? this.state.delegateToIdInSubSite : this.state.dccreviewer),
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
                  Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + "&wf=dcc"
                },
              });

              //MY tasks list updation
              await this.reqWeb.lists.getByTitle(this.props.WorkflowTasksList).items.add
                ({
                  Title: "Document Controller Review '" + this.state.documentName + "'",
                  Description: "DCC Review request for  '" + this.state.documentName + "' by '" + this.state.currentuser + "' on '" + this.today + "'",
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
                    Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + "&wf=dcc"
                  },

                }).then(async taskId => {
                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                    ({
                      TaskID: taskId.data.ID,
                    });
                  
                  //notification preference checking                                 
                   this._sendmail(this.state.dccreviewerEmail,"DocDCCReview",this.state.dccreviewerName)                                            
                   .then(aftermail=>{
                     //Email pending  emailbody to approver                 
                       this.validator.hideMessages();                   
                  window.location.replace(this.RedirectUrl);
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
    }).then(async update =>{
 await sp.web.getList(this.props.siteUrl + "/Lists/"+this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus:"Under Review",
        Workflow:"Review"
      });
      await sp.web.getList(this.props.siteUrl + "/"+this.props.SourceDocumentLibrary).items.getById(this.SourceDocumentID).update({
        WorkflowStatus:"Under Review",
        Workflow:"Review"
      });
     await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
        Title:this.state.documentID,
        Status:"Under Review ",
        LogDate:this.today,
        WorkflowID:this.newheaderid,
        Revision:this.state.revision,
        Workflow:"DCC Review",
        DocumentIndexId:this.documentIndexID
      });
    }).then(msg=>{
      this.setState({
        statusMessage: { isShowMessage: true, message: this.DccReview, messageType: 4 },
      });
  });//msg
 });//newheaderid
  }
  public async underApprove(previousHeaderItem) {

    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Approval",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.SourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      // ReviewersId: { results: this.CurrentId },
      RequesterId: this.CurrentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Approve",
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }).then(async i => {
      this.newheaderid = i.data.ID;
       //Task delegation getting user id from hubsite
    this.reqWeb.siteUsers.getByEmail(this.state.approverEmail).get().then(async user => {
      console.log('User Id: ', user.Id);
      this.setState({
        hubSiteUserId: user.Id,
      });
      //Task delegation 
      const taskDelegation: any[] = await this.reqWeb.lists.getByTitle(this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "'").get();
      console.log(taskDelegation);
      if (taskDelegation.length > 0) {
        let duedate = moment(this.state.dueDate).toDate();
        let ToDate = moment(taskDelegation[0].ToDate).toDate();
        let FromDate = moment(taskDelegation[0].FromDate).toDate();
        duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
        ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
        FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
        if (moment(duedate).isBetween(FromDate, ToDate) ||moment(duedate).isSame(FromDate)||moment(duedate).isSame(ToDate)) {
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
                  Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentApprovalPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                },
              });

              //MY tasks list updation
              await this.reqWeb.lists.getByTitle(this.props.WorkflowTasksList).items.add
                ({
                  Title: "Approve '" + this.state.documentName + "'",
                  Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentuser + "' on '" + this.today + "'",
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
                    Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentApprovalPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                  },

                }).then(taskId => {
                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                    ({
                      TaskID: taskId.data.ID,
                    });
                  //notification preference checking                                 
                   this._sendmail(this.state.approverEmail,"DocApproval",this.state.approverName)                                            
                   .then(aftermail=>{
                     //Email pending  emailbody to approver                 
                       this.validator.hideMessages();                   
                  window.location.replace(this.RedirectUrl);
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
    }).then(async update =>{
 await sp.web.getList(this.props.siteUrl + "/Lists/"+this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus:"Under Approval",
        Workflow:"Approval"
      });
      await sp.web.getList(this.props.siteUrl + "/"+this.props.SourceDocumentLibrary).items.getById(this.SourceDocumentID).update({
        WorkflowStatus:"Under Approval",
        Workflow:"Approval"
      });
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
        Title:this.state.documentID,
        Status:"Under Approval",
        LogDate:this.today,
        WorkflowID:this.newheaderid,
        Revision:this.state.revision,
        DocumentIndexId:this.documentIndexID,
        Workflow:"Approve"
      });
    }).then(msg=>{
      this.setState({
        statusMessage: { isShowMessage: true, message: this.UnderApproval, messageType: 4 },
      });
  });//msg
 });//newheaderid
  }
  public async underReview(previousHeaderItem) {
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.SourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ReviewersId: { results: this.state.reviewers },
      ApproverId: this.state.approver,
      RequesterId: this.CurrentId,
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
            const taskDelegation: any[] = await this.reqWeb.lists.getByTitle(this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + hubsieUser.Id + "'").get();
            console.log(taskDelegation);
            //Check if Task Delegation
             if (taskDelegation.length > 0) {
              let duedate = moment(this.state.dueDate).toDate();
              let ToDate = moment(taskDelegation[0].ToDate).toDate();
              let FromDate = moment(taskDelegation[0].FromDate).toDate();
              duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
              ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
              FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
              if (moment(duedate).isBetween(FromDate, ToDate) ||moment(duedate).isSame(FromDate)||moment(duedate).isSame(ToDate)) {
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
                          Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                        },
                      });//Update link
                  //MY tasks list updation with delegated from
                      await this.reqWeb.lists.getByTitle(this.props.WorkflowTasksList).items.add({
                          Title: "Review '" + this.state.documentName + "'",
                          Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentuser + "' on '" + this.today + "'",
                          DueDate: this.state.dueDate,
                          StartDate: this.today,
                          AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : hubsieUser.Id),
                          Priority: (this.state.CriticalDocument == true ? "Critical" : ""),
                          DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                          Source: (this.props.project ? "Project" : "QDMS"),
                          DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : parseInt("")),
                          Workflow: "Review",
                          Link: {
                            "__metadata": { type: "SP.FieldUrlValue" },
                            Description: "Link to Review",
                            Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                          },
                        }).then(taskId => {
                          sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                            ({
                              TaskID: taskId.data.ID,
                            });
                          this._sendmail(DelegatedTo.Email,"DocReview",DelegatedTo.Title);  
                          this.validator.hideMessages();                   
                          window.location.replace(this.RedirectUrl);
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
            else{
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
                          Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                        },
                      });
                     //MY tasks list updation with delegated from
                      await this.reqWeb.lists.getByTitle(this.props.WorkflowTasksList).items.add({
                          Title: "Review '" + this.state.documentName + "'",
                          Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentuser + "' on '" + this.today + "'",
                          DueDate: this.state.dueDate,
                          StartDate: this.today,
                          AssignedToId:  hubsieUser.Id,
                          Priority: (this.state.CriticalDocument == true ? "Critical" : ""),
                          Source: (this.props.project ? "Project" : "QDMS"),
                           Workflow: "Review",
                          Link: {
                            "__metadata": { type: "SP.FieldUrlValue" },
                            Description: "Link to Review",
                            Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                          },
                        }).then(taskId => {
                          sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                            ({
                              TaskID: taskId.data.ID,
                            });
                          this._sendmail(user.Email,"DocReview",user.Title);  
                          this.validator.hideMessages();                   
                          window.location.replace(this.RedirectUrl);
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
      await sp.web.getList(this.props.siteUrl + "/Lists/"+this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus:"Under Review",
        Workflow:"Review"
      });
      await sp.web.getList(this.props.siteUrl + "/"+this.props.SourceDocumentLibrary).items.getById(this.SourceDocumentID).update({
        WorkflowStatus:"Under Review",
        Workflow:"Review"
      });
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
        Title:this.state.documentID,
        Status:"Under Review",
        LogDate:this.today,
        WorkflowID:this.newheaderid,
        Revision:this.state.revision,
        DocumentIndexId:this.documentIndexID,
        Workflow:"Review"
      });
    }
    });
  }
  public async underProjectApprove(previousHeaderItem) {

    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Approval",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.SourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ApproverId: this.state.approver,
      ReviewersId: { results: this.state.reviewers },
      RequesterId: this.CurrentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Approve",
      DocumentControllerId:this.state.dccreviewer,
      RevisionLevelId:this.state.RevisionLevelvalue,
      RevisionCodingId:this.state.RevisionCoding,
      OwnerId: this.state.ownerId,
      PreviousReviewHeader: previousHeaderItem.toString()
    }).then(async i => {
      this.newheaderid = i.data.ID;
       //Task delegation getting user id from hubsite
    this.reqWeb.siteUsers.getByEmail(this.state.approverEmail).get().then(async user => {
      console.log('User Id: ', user.Id);
      this.setState({
        hubSiteUserId: user.Id,
      });
      //Task delegation 
      const taskDelegation: any[] = await this.reqWeb.lists.getByTitle(this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + user.Id + "'").get();
      console.log(taskDelegation);
      if (taskDelegation.length > 0) {
        let duedate = moment(this.state.dueDate).toDate();
        let ToDate = moment(taskDelegation[0].ToDate).toDate();
        let FromDate = moment(taskDelegation[0].FromDate).toDate();
        duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
        ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
        FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
        if (moment(duedate).isBetween(FromDate, ToDate) ||moment(duedate).isSame(FromDate)||moment(duedate).isSame(ToDate)) {
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
                  Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentApprovalPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                },
              });

              //MY tasks list updation
              await this.reqWeb.lists.getByTitle(this.props.WorkflowTasksList).items.add
                ({
                  Title: "Approve '" + this.state.documentName + "'",
                  Description: "Approval request for  '" + this.state.documentName + "' by '" + this.state.currentuser + "' on '" + this.today + "'",
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
                    Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentApprovalPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                  },

                }).then(taskId => {
                  sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                    ({
                      TaskID: taskId.data.ID,
                    });
                  //notification preference checking                                 
                   this._sendmail(this.state.approverEmail,"DocApproval",this.state.approverName)                                            
                   .then(aftermail=>{
                     //Email pending  emailbody to approver                 
                       this.validator.hideMessages();                   
                  window.location.replace(this.RedirectUrl);
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
    }).then(async update =>{
 await sp.web.getList(this.props.siteUrl + "/Lists/"+this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus:"Under Approval",
        Workflow:"Approval"
      });
      await sp.web.getList(this.props.siteUrl + "/"+this.props.SourceDocumentLibrary).items.getById(this.SourceDocumentID).update({
        WorkflowStatus:"Under Approval",
        Workflow:"Approval"
      });
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
        Title:this.state.documentID,
        Status:"Under Approval",
        LogDate:this.today,
        WorkflowID:this.newheaderid,
        Revision:this.state.revision,
        DocumentIndexId:this.documentIndexID,
        Workflow:"Approve"
      });
    }).then(msg=>{
      this.setState({
        statusMessage: { isShowMessage: true, message: this.UnderApproval, messageType: 4 },
      });
  });//msg
 });//newheaderid
  }
  public async underProjectReview(previousHeaderItem) {
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowHeaderList).items.add({
      Title: this.state.documentName,
      DocumentID: this.state.documentID,
      WorkflowStatus: "Under Review",
      Revision: this.state.revision,
      DueDate: this.state.dueDate,
      SourceDocumentID: this.SourceDocumentID,
      DocumentIndexId: this.documentIndexID,
      DocumentIndexID: this.documentIndexID,
      ReviewersId: { results: this.state.reviewers },
      ApproverId: this.state.approver,
      RequesterId: this.CurrentId,
      RequesterComment: this.state.comments,
      RequestedDate: this.today,
      Workflow: "Review",
      DocumentControllerId:this.state.dccreviewer,
      RevisionLevelId:this.state.RevisionLevelvalue,
      RevisionCodingId:this.state.RevisionCoding,
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
            const taskDelegation: any[] = await this.reqWeb.lists.getByTitle(this.props.TaskDelegationSettings).items.select("DelegatedFor/ID,DelegatedFor/Title,DelegatedFor/EMail,DelegatedTo/ID,DelegatedTo/Title,DelegatedTo/EMail,FromDate,ToDate").expand("DelegatedFor,DelegatedTo").filter("DelegatedFor/ID eq '" + hubsieUser.Id + "'").get();
            console.log(taskDelegation);
            //Check if Task Delegation
             if (taskDelegation.length > 0) {
              let duedate = moment(this.state.dueDate).toDate();
              let ToDate = moment(taskDelegation[0].ToDate).toDate();
              let FromDate = moment(taskDelegation[0].FromDate).toDate();
              duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
              ToDate = new Date(ToDate.getFullYear(), ToDate.getMonth(), ToDate.getDate());
              FromDate = new Date(FromDate.getFullYear(), FromDate.getMonth(), FromDate.getDate());
              if (moment(duedate).isBetween(FromDate, ToDate) ||moment(duedate).isSame(FromDate)||moment(duedate).isSame(ToDate)) {
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
                          Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                        },
                      });//Update link
                  //MY tasks list updation with delegated from
                      await this.reqWeb.lists.getByTitle(this.props.WorkflowTasksList).items.add({
                          Title: "Review '" + this.state.documentName + "'",
                          Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentuser + "' on '" + this.today + "'",
                          DueDate: this.state.dueDate,
                          StartDate: this.today,
                          AssignedToId: (this.state.delegatedToId != "" ? this.state.delegatedToId : hubsieUser.Id),
                          Priority: (this.state.CriticalDocument == true ? "Critical" : ""),
                          DelegatedOn: (this.state.delegatedToId !== "" ? this.today : " "),
                          Source: (this.props.project ? "Project" : "QDMS"),
                          DelegatedFromId: (this.state.delegatedToId != "" ? this.state.delegatedFromId : parseInt("")),
                          Workflow: "Review",
                          Link: {
                            "__metadata": { type: "SP.FieldUrlValue" },
                            Description: "Link to Review",
                            Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                          },
                        }).then(taskId => {
                          sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                            ({
                              TaskID: taskId.data.ID,
                            });
                          this._sendmail(DelegatedTo.Email,"DocReview",DelegatedTo.Title);  
                          this.validator.hideMessages();                   
                          window.location.replace(this.RedirectUrl);
                          this.setState({
                            statusMessage: { isShowMessage: true, message:this.UnderReview, messageType: 4 },
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
            else{
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
                          Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                        },
                      });
                     //MY tasks list updation with delegated from
                      await this.reqWeb.lists.getByTitle(this.props.WorkflowTasksList).items.add({
                          Title: "Review '" + this.state.documentName + "'",
                          Description: "Review request for  '" + this.state.documentName + "' by '" + this.state.currentuser + "' on '" + this.today + "'",
                          DueDate: this.state.dueDate,
                          StartDate: this.today,
                          AssignedToId:  hubsieUser.Id,
                          Priority: (this.state.CriticalDocument == true ? "Critical" : ""),
                          Source: (this.props.project ? "Project" : "QDMS"),
                           Workflow: "Review",
                          Link: {
                            "__metadata": { type: "SP.FieldUrlValue" },
                            Description: "Link to Review",
                            Url: this.props.siteUrl + "/SitePages/"+this.props.DocumentReviewPage+".aspx?hid=" + this.newheaderid + "&dtlid=" + r.data.ID + ""
                          },
                        }).then(taskId => {
                          sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.WorkflowDetailsList).items.getById(r.data.ID).update
                            ({
                              TaskID: taskId.data.ID,
                            });
                          this._sendmail(user.Email,"DocReview",user.Title);  
                          this.validator.hideMessages();                   
                          window.location.replace(this.RedirectUrl);
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
      await sp.web.getList(this.props.siteUrl + "/Lists/"+this.props.DocumentIndexList).items.getById(this.documentIndexID).update({
        WorkflowStatus:"Under Review",
        Workflow:"Review"
      });
      await sp.web.getList(this.props.siteUrl + "/"+this.props.SourceDocumentLibrary).items.getById(this.SourceDocumentID).update({
        WorkflowStatus:"Under Review",
        Workflow:"Review"
      });
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.DocumentRevisionLogList).items.add({
        Title:this.state.documentID,
        Status:"Under Review",
        LogDate:this.today,
        WorkflowID:this.newheaderid,
        Revision:this.state.revision,
        DocumentIndexId:this.documentIndexID,
        Workflow:"Review"
      });
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
          <div className={styles.flex}>
            <div className={styles.width}><Label >Document : <a href={this.state.linkToDoc}>{this.state.documentName}</a></Label></div>
            <div ><Label >Revision : {this.state.revision}</Label></div>
          </div>
          <div className={styles.flex}>
            <div className={styles.width}><Label >Owner : {this.state.ownerName} </Label></div>
            <div><Label >Requester : {this.state.currentuser}</Label></div>
          </div>
          <div hidden={this.state.hideproject}>
            <div className={styles.flex} >
              <div className={styles.width} style={{paddingRight:'15px'}}>
                <Dropdown
                  placeholder="Select Option"
                  label="RevisionLevel"
                  style={{ marginBottom: '10px', backgroundColor: "white", height:'34px'}}
                  options={this.state.RevisionLevelArray}
                  onChanged={this.RevisionLevelChanged}
                  selectedKey={this.state.RevisionLevelvalue}
                  required />
                  <div style={{ color: "#dc3545" }}>{this.validator.message("RevisionLevel", this.state.RevisionLevelvalue, "required")}{" "}</div>
                  </div>
<div className ={styles.width}>
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
            selectedItems={(items) => this._ReviewerChange(items)}
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
                selectedItems={(items) => this._ApproverChange(items)}
                //defaultSelectedUsers={[this.state.approver]}
                showHiddenInUI={false}
                isRequired={true}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              /> <div style={{ color: "#dc3545" }}>{this.validator.message("Approver", this.state.approver, "required")}{" "}</div>
            </div>
            <div className ={styles.width} style={{paddingLeft:'15px',marginTop:'2px'}}>
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
              <DefaultButton id="b2" className={styles.btn} disabled={this.state.savedisable} onClick={this._submitSendRequest}>Submit</DefaultButton >
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
          
         
            <div className={styles.width}><Label >Document : <a href={this.state.linkToDoc}>{this.state.documentName}</a></Label></div>
            <div ><Label >Revision : {this.state.revision}</Label></div>
          
         
            <div className={styles.width}><Label >Owner : {this.state.ownerName} </Label></div>
            <div><Label >Requester : {this.state.currentuser}</Label></div>
          
          <div hidden={this.state.hideproject}>
            
          <div className={styles.drpdwn} >
                <Dropdown
                  placeholder="Select Option"
                  label="RevisionLevel"
                  style={{ marginBottom: '10px', backgroundColor: "white" }}
                  options={this.state.RevisionLevelArray}
                  onChanged={this.RevisionLevelChanged}
                  selectedKey={this.state.RevisionLevelvalue}
                  required />
                  <div style={{ color: "#dc3545" }}>{this.validator.message("RevisionLevel", this.state.RevisionLevelvalue, "required")}{" "}</div>
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
            selectedItems={(items) => this._ReviewerChange(items)}
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
                selectedItems={(items) => this._ApproverChange(items)}
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
              <DefaultButton id="b2" className={styles.btn} disabled={this.state.savedisable} onClick={this._submitSendRequest}>Submit</DefaultButton >
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
