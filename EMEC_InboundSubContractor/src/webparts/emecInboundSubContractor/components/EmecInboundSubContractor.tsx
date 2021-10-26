import * as React from 'react';
import styles from './EmecInboundSubContractor.module.scss';
import { IEmecInboundSubContractorProps } from './IEmecInboundSubContractorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, IconButton, IIconProps, Label, TextField } from '@microsoft/office-ui-fabric-react-bundle';
import { Checkbox, DatePicker, DialogFooter, Dropdown, IDropdownOption, SearchBox } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp, IList, Web } from "@pnp/sp/presets/all";
import * as _ from 'lodash';
import Select from 'react-select-plus';
import 'react-select-plus/dist/react-select-plus.css';
import * as moment from 'moment';
export interface IEmecInboundSubContractorState {
  dcc: any;
  owner: any;
  recievedDate: Date;
  subContractorNumber: any;
  poNumber: any;
  comments: any;
  projectName: any;
  projectNumber: any;
  revisionSettingsArray: any[];
  transmittalSettingsArray: any[];
  subContractorArray: any[];
  subContractorkey: any;
  multidealer: boolean;
  transmittalOutlookDocumentArray: any[];
  documentIndexArray:any[];
  documentIndexID:any;
  documentIndexTitle:any;
  revisionCodingId:any;
  isIncrement:boolean;
  transmittalOutlookId:any;
  transmittalKey:any;
  ownerId:any;
  showGrid:boolean;
  gridDocument:any[];
  externalDate:Date;
  externalComments:any;
  gridExternalDocument:any[];
  showExternalGrid:boolean;
  incrementSequenceNumber:any;
  transmittalSubContractorId:any;

}
export default class EmecInboundSubContractor extends React.Component<IEmecInboundSubContractorProps, IEmecInboundSubContractorState, {}> {
  private reqWeb = Web(this.props.hubUrl);
  private currentEmail;
  private currentUserTitle;
  private currentUserId;
  private addDocument = [];
  private addExternalDocument = [];
  public constructor(props: IEmecInboundSubContractorProps) {
    super(props);
    this.state = {
      dcc: "",
      owner: "",
      recievedDate: null,
      subContractorNumber: "",
      poNumber: "",
      comments: "",
      projectName: "",
      projectNumber: "",
      revisionSettingsArray: [],
      transmittalSettingsArray: [],
      subContractorArray: [],
      subContractorkey: "",
      multidealer: false,
      transmittalOutlookDocumentArray: [],
      documentIndexArray:[],
      documentIndexID:"",
      documentIndexTitle:"",
      revisionCodingId:"",
      isIncrement:false,
      transmittalOutlookId:"",
      transmittalKey:"",
      ownerId:"",
      showGrid:true,
      gridDocument:[],
      externalDate:null,
      externalComments:"",
      gridExternalDocument:[],
      showExternalGrid:true,
      incrementSequenceNumber:"",
      transmittalSubContractorId:""
    };
    this._bindData = this._bindData.bind(this);
    this._getSubContractor = this._getSubContractor.bind(this);
    this._getProjectInformation = this._getProjectInformation.bind(this);
    this._getRevisionSettings = this._getRevisionSettings.bind(this);
    this._getTransmittalSettings = this._getTransmittalSettings.bind(this);
    this._subContactorChanged = this._subContactorChanged.bind(this);
    this._getDocumentIndex = this._getDocumentIndex.bind(this);
    this._documentIndexChange =this._documentIndexChange.bind(this);
    this._outlookDocumentChange = this._outlookDocumentChange.bind(this);
    this._dccChange = this._dccChange.bind(this);
    this._ownerChange = this._ownerChange.bind(this);
    this._onRecievedDatePickerChange = this._onRecievedDatePickerChange.bind(this);
    this._subContractorNumberChange =this._subContractorNumberChange.bind(this);
    this._poNumberChange = this._poNumberChange.bind(this);
    this._onIncrementRevisionChecked = this._onIncrementRevisionChecked.bind(this);
    this._commentschange = this._commentschange.bind(this);
    this._addindex =this._addindex.bind(this);
    this._onTransmittalSettingsChange =this._onTransmittalSettingsChange.bind(this);
    this._onRevisionCodingChange = this._onRevisionCodingChange.bind(this);
    this._onDatePickerChange = this._onDatePickerChange.bind(this);
    this._externalCommentsChange = this._externalCommentsChange.bind(this);
    this._addexternalindex = this._addexternalindex.bind(this);
    this._saveAsDraft = this._saveAsDraft.bind(this);
    this._idGeneration = this._idGeneration.bind(this);
    this._transmittalSequenceNumber = this._transmittalSequenceNumber.bind(this);




  }
  public async componentDidMount() {
    //Get Current User
    const user = await sp.web.currentUser.get();
    this.currentEmail = user.Email;  
    this.currentUserTitle = user.Title;
    this.currentUserId = user.Id;
    let getdccreviewer = [];
    getdccreviewer.push(this.currentUserTitle);
    this.setState({
      dcc: getdccreviewer[0]
    });
    let today = new Date();
    this.setState({
      recievedDate:today,
      externalDate:today
    });

    this._bindData();
  }
  public async _bindData() {
    this._getSubContractor();
    this._getProjectInformation();
    this._getRevisionSettings();
    this._getTransmittalSettings();
    this._getDocumentIndex();

    

  }
  public async _getSubContractor() {
    let subContractorarray = [];
    let sorted_SubContractor = [];
    let subContractor;
    const subContractoritems: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.companyList).items.select("Title,ID,Active").filter("CompanyType eq 'Sub-Contractor' ").get();
    for (let i = 0; i < subContractoritems.length; i++) {
      if (subContractoritems[i].Active == true) {
        subContractor = {
          key: subContractoritems[i].ID,
          text: subContractoritems[i].Title
        };
      }
      subContractorarray.push(subContractor);
    }
    sorted_SubContractor = _.orderBy(subContractorarray, 'text', ['asc']);
    this.setState({
      subContractorArray: sorted_SubContractor
    });

  }
  public async _getProjectInformation() {
    const projectInformation = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.projectInformationListName).items.get();
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
  public async _getRevisionSettings() {
    let revisionSettingsArray = [];
    let sorted_RevisionSettings = [];
    const revisionSettingsItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.revisionLevelList).items.select("Title,ID").get();
    for (let i = 0; i < revisionSettingsItem.length; i++) {
      let revisionSettingsItemdata = {
        key: revisionSettingsItem[i].ID,
        text: revisionSettingsItem[i].Title
      };
      revisionSettingsArray.push(revisionSettingsItemdata);

    }
    
    sorted_RevisionSettings = _.orderBy(revisionSettingsArray, 'text', ['asc']);
    this.setState({
      revisionSettingsArray: sorted_RevisionSettings
    });
  }
  public async _getTransmittalSettings() {
    let transmittalCodeSettingsArray = [];
    let sorted_transmittalCodeSettings = [];
    const transmittalCodeSettingsItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.transmittalCodeSettings).items.get();

    
    for (let i = 0; i < transmittalCodeSettingsItem.length; i++) {
      if(transmittalCodeSettingsItem[i].AcceptanceCode == false){
      let transmittalCodeSettingsItemdata = {
        key: transmittalCodeSettingsItem[i].ID,
        text: transmittalCodeSettingsItem[i].Title
      };
      transmittalCodeSettingsArray.push(transmittalCodeSettingsItemdata);
    }
    }
    
    sorted_transmittalCodeSettings = _.orderBy(transmittalCodeSettingsArray, 'text', ['asc']);
    this.setState({
      transmittalSettingsArray: sorted_transmittalCodeSettings
    });
  }
  public async _getDocumentIndex() {
    let documentIndexArray = [];
    let sorted_documentIndexArray = [];
    const documentIndexArrayItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.get();
   
    for (let i = 0; i < documentIndexArrayItem.length; i++) {
      if(documentIndexArrayItem[i].ExternalDocument == true){
      let documentIndexArrayItemdata = {
        key: documentIndexArrayItem[i].ID,
        text: documentIndexArrayItem[i].DocumentName
      };
      documentIndexArray.push(documentIndexArrayItemdata);
    }
    }
   
    sorted_documentIndexArray = _.orderBy(documentIndexArray, 'text', ['asc']);
    this.setState({
      documentIndexArray: sorted_documentIndexArray
    });
  }
  public async _subContactorChanged(option: { key: any; text: any }) {
    let transmittalOutlookDocumentArray = [];
    let sorted_transmittalOutlookDocumentArray = [];
    this.setState({ subContractorkey: option.text });
    const document = await sp.web.getList(this.props.siteUrl + "/" + this.props.transmittalOutlookLibrary).items.filter("From eq 'Sub-Contractor'").select("ID,BaseName,SubContractor").get();
    
    for (let i = 0; i < document.length; i++) {
      if (document[i].SubContractor == option.text) {
        let transmittalOutlookDocument = {
          key: document[i].ID,
          text: document[i].BaseName
        };
        transmittalOutlookDocumentArray.push(transmittalOutlookDocument);
      }
    }
    sorted_transmittalOutlookDocumentArray = _.orderBy(transmittalOutlookDocumentArray, 'text', ['asc']);
    this.setState({
      transmittalOutlookDocumentArray: sorted_transmittalOutlookDocumentArray
    });



  }
  public async _documentIndexChange(option: { key: any; text: any }) { 
    
    const documentIndexItem = await  sp.web.getList(this.props.siteUrl+"/Lists/"+this.props.documentIndexList).items.select("Owner/Title,Owner/ID,RevisionCoding/Title,RevisionCoding/ID").expand("Owner,RevisionCoding").filter("ID eq '"+option.key+"'").get();
    
      this.setState({
         documentIndexID: option.key,
         documentIndexTitle:option.text,
         owner:documentIndexItem[0].Owner.Title,
         ownerId:documentIndexItem[0].Owner.ID,
         revisionCodingId:documentIndexItem[0].RevisionCoding.ID
         });
   
    
  }
  public async _outlookDocumentChange(option: { key: any; text: any }) {
    const document = await sp.web.getList(this.props.siteUrl + "/" + this.props.transmittalOutlookLibrary).items.getById(option.key).get();
    this.setState({
      subContractorNumber:document.SubContractorDocumentId,
      poNumber:document.PONumber,
      transmittalOutlookId:option.key
      });
   }
  public _dccChange = (items: any[]) => {
    
    let getSelecteddccreviewer = [];

    for (let item in items) {

      getSelecteddccreviewer.push(items[item].id);
    }
    this.setState({
      dcc: getSelecteddccreviewer[0]
    });
  }
  public _ownerChange = (items: any[]) => {
   
    let getSelectedOwner = [];

    for (let item in items) {

      getSelectedOwner.push(items[item].id);
    }
    this.setState({
      ownerId: getSelectedOwner[0]
    });
  }
  private _onRecievedDatePickerChange = (date?: Date): void => {
    this.setState({ recievedDate: date });
  }
  public _subContractorNumberChange= (ev: React.FormEvent<HTMLInputElement>, subContractorNumber?: string) => {
    
      this.setState({ subContractorNumber: subContractorNumber || '' });
  
   }
  public _poNumberChange= (ev: React.FormEvent<HTMLInputElement>, poNumber?: string) => {
    this.setState({ poNumber: poNumber || '' });
   }
  public _onTransmittalSettingsChange(option: { key: any; text: any }) {
    this.setState({ transmittalKey: option.key });
  }
  public _onRevisionCodingChange(option: { key: any; text: any }) {
    this.setState({ revisionCodingId: option.key });
  }
  private _onIncrementRevisionChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ isIncrement: true});
     }
     else{
      this.setState({ isIncrement: false});
     }
    }
  //Comment Change
  public _commentschange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ comments: comments });
  }
  public async _addindex() {
    
    if ((document.querySelector("#newfile") as HTMLInputElement).files[0] != null) {
      let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
      var docname = myfile.name;
      alert(docname);
      if (myfile.size <= 10485760) {
        sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/InboundTemperory").files.add(docname, myfile, true)
        .then(fileUploaded => {
            console.log("File Uploaded");
            fileUploaded.file.getItem().then(item=> {
                let DocId=item["ID"];
                item.update({ 
                  SubContractor: this.state.subContractorkey,
                  DocumentControllerId: this.currentUserId,
                  // TransmittalOutlookLibraryId:this.state.transmittalOutlookId,
                  DocumentIndexId:this.state.documentIndexID,
                  OwnerId:this.state.ownerId,
                  SubContractorDocId:this.state.subContractorNumber,
                  ReceivedDate:this.state.recievedDate,
                  PONumber:this.state.poNumber,
                  TransmittalCodeId:this.state.transmittalKey,
                  RevisionCodeId:this.state.revisionCodingId,
                  IncrementRevision:this.state.isIncrement,
                  Comments:this.state.comments
                 });
                 this.addDocument.push({
                  SubContractor: this.state.subContractorkey,
                  DocumentControllerId: this.currentUserId,
                  // TransmittalOutlookLibraryId:this.state.transmittalOutlookId,
                  DocumentIndexId:this.state.documentIndexID,
                  DocumentIndex:this.state.documentIndexTitle,
                  OwnerId:this.state.ownerId,
                  Owner:this.state.owner,
                  SubContractorDocId:this.state.subContractorNumber,
                  ReceivedDate:moment(this.state.recievedDate).format("DD/MM/YYYY hh:mm a"),
                  PONumber:this.state.poNumber,
                  TransmittalCodeId:this.state.transmittalKey,
                  RevisionCodeId:this.state.revisionCodingId,
                  IncrementRevision:this.state.isIncrement,
                  Comments:this.state.comments
              });
              console.log(this.addDocument);
            this.setState({
                            gridDocument:this.addDocument,
                            showGrid:false
                        });
             });
          });
        }
        }
        else if(this.state.transmittalOutlookId != "" ){
          const document = await sp.web.getList(this.props.siteUrl + "/" + this.props.transmittalOutlookLibrary).items.getById(parseInt(this.state.transmittalOutlookId)).get();
         console.log(document);
         sp.web.getFileByServerRelativeUrl(this.props.siteUrl + "/"+this.props.transmittalOutlookLibrary+"/" + this.state.transmittalOutlookId.toString()).getItem()
         .then(templateData => {
            //  return sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/SourceDocuments/")
            //      .files.add(this.state.docExtension, templateData, true);
         }).then(f => {
            //  console.log("File Uploaded");
            //  f.file.getItem().then(item => {
            //      let sourceDocId=item["ID"].toNumber;
            //  });
            });
          this.setState({
           
            });

        }
        else{
          alert("no document");
        }

   }
   public itemDeleteFromGrid(items,key){

   }
   public _onDatePickerChange= (date?: Date): void => {
    this.setState({ externalDate: date });
  }
   public _externalCommentsChange= (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ externalComments: comments });
   }
   public _addexternalindex(){
    if ((document.querySelector("#externalFile") as HTMLInputElement).files[0] != null) {
      let myfile = (document.querySelector("#externalFile") as HTMLInputElement).files[0];
      var docname = myfile.name;
      alert(docname);
      this.addExternalDocument.push({
        Name: myfile,
        DocName: docname,
        ExternalDate:moment(this.state.externalDate).format("DD/MM/YYYY hh:mm a"),
        Comments:this.state.externalComments
    });
    console.log(this.addExternalDocument);
  this.setState({
                  gridExternalDocument:this.addExternalDocument,
                  showExternalGrid:false
              });
        }
        
        else{
          alert("no document");
        }
   }
  public async _saveAsDraft(){
    this._idGeneration();
   }
   public _idGeneration(){
    let prefix;
    let separator;
    let sequenceNumber;
    let title;
    let counter;
    let transmittalSubContractorId;
   
    sp.web.getList(this.props.siteUrl+"/Lists/"+this.props.transmittalIdSettings).items.filter("TransmittalCategory eq 'Inbound Sub-contractor'").get().then(transmittalIdSettingsItems =>{
      console.log("transmittalIdSettingsItems",transmittalIdSettingsItems);
      prefix = transmittalIdSettingsItems[0].Prefix;
      separator = transmittalIdSettingsItems[0].Separator;
      sequenceNumber = transmittalIdSettingsItems[0].SequenceNumber;
      title = transmittalIdSettingsItems[0].Title;
      counter = transmittalIdSettingsItems[0].Counter;
      let increment =counter+1;
      var incrementvalue =increment.toString();
      this._transmittalSequenceNumber(incrementvalue,sequenceNumber);
      transmittalSubContractorId = prefix+separator+title+separator+this.state.projectNumber+separator+this.state.incrementSequenceNumber;
      console.log("transmittalID",transmittalSubContractorId);
      this.setState({
        transmittalSubContractorId :  transmittalSubContractorId,
       });
    });
   }
   private _transmittalSequenceNumber(incrementvalue, sequenceNumber) {
    var incrementSequenceNumber = incrementvalue;
    while (incrementSequenceNumber.length < sequenceNumber)
    incrementSequenceNumber = "0" + incrementSequenceNumber;
     console.log(incrementSequenceNumber);
     this.setState({
      incrementSequenceNumber :  incrementSequenceNumber,
     });
  }
  public _submit(){ }
  public _onCancel(){ }


  public render(): React.ReactElement<IEmecInboundSubContractorProps> {
    const DeleteIcon: IIconProps = { iconName: 'Delete' };
    const AddIcon: IIconProps = { iconName: 'CircleAdditionSolid' };
   
    return (
      <div className={styles.emecInboundSubContractor}>
        <div className={styles.alignCenter}>Inbound Transmittal  </div>
        <div className={styles.divrow}>
          <div className={styles.wdthrgt}><Label>Transmittal ID : TRM-IB-0001 </Label></div>
          <div className={styles.wdthlft}><Label>Transmittal Date : 16 Aug 2021</Label></div>
        </div>
        <div className={styles.divrow}>
          <div className={styles.wdthrgt}><Label >Project Name : {this.state.projectName} </Label></div>
          <div className={styles.wdthlft}><Label >Project Number :{this.state.projectNumber} </Label></div>
        </div>

        <div className={styles.divrow}>
          <div className={styles.wdthrgt}>
            {/* <Label >Sub-Contractor : </Label> */}
            <Dropdown
              placeholder="Sub-Contractor:"
              label="Select Sub-Contractor"
              options={this.state.subContractorArray}
              onChanged={this._subContactorChanged}
            />

          </div>
          <div className={styles.wdthlft}>
            <PeoplePicker
              context={this.props.context}
              titleText="DCC"
              personSelectionLimit={1}
              groupName={""} // Leave this blank in case you want to filter from all users    
              showtooltip={true}
              disabled={false}
              ensureUser={true}
              onChange={(items) => this._dccChange(items)}
              defaultSelectedUsers={[this.state.dcc]}
              showHiddenInUI={false}
              // isRequired={true}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            />
          </div>
        </div>
        <hr />
        <Label>Project Document </Label>
        <div className={styles.divrow}>
          <div className={styles.wdthrgt}>
            <Label >Upload Document:</Label>
            <input type="file" name="myFile" id="newfile"></input>
          </div>
          <div className={styles.wdthlft} >
            <Dropdown
              placeholder="Search Document"
              label="Select Document"
              options={this.state.transmittalOutlookDocumentArray}
              onChanged={this._outlookDocumentChange}
            />
          </div>
        </div>
        <div className={styles.divrow}>
          <div className={styles.wdthrgt}>
            <Dropdown placeholder="Select Document Index" 
                      label="Document Index"
                      options={this.state.documentIndexArray}
                      onChanged={this._documentIndexChange}
            />
          </div>
          <div className={styles.wdthlft}>
            <PeoplePicker
              context={this.props.context}
              titleText="Owner"
              personSelectionLimit={1}
              groupName={""} // Leave this blank in case you want to filter from all users    
              showtooltip={true}
              disabled={false}
              ensureUser={true}
              onChange={(items) => this._ownerChange(items)}
              defaultSelectedUsers={[this.state.owner]}
              showHiddenInUI={false}
              // isRequired={true}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
            />
          </div>
        </div>
        <div className={styles.divrow}>
          <div className={styles.wdthrgt}>
            <TextField label="SubContractor Doc No" onChange={this._subContractorNumberChange} value={this.state.subContractorNumber}></TextField>
            </div>
          <div className={styles.wdthlft}>
            <DatePicker label="Recieved Date" value={this.state.recievedDate} onSelectDate={this._onRecievedDatePickerChange} placeholder="Select a date" />
            </div>
        </div>
        <div className={styles.divrow}>
          <div className={styles.wdthrgt}>
            <TextField label="PO Number" onChange={this._poNumberChange} value={this.state.poNumber}></TextField>
            </div>
          <div className={styles.wdthlft}>
            <Dropdown placeholder="Select Transmittal Code" label="Transmittal Code" options={this.state.transmittalSettingsArray} onChanged={this._onTransmittalSettingsChange}/>
            </div>
        </div>
        <div className={styles.divrow}>
          <div className={styles.wdthrgt}>
            <Dropdown placeholder="Select Revision Code" label="Revision Code" options={this.state.revisionSettingsArray} selectedKey={this.state.revisionCodingId} onChanged={this._onRevisionCodingChange}/>
            </div>
          <div className={styles.wdthlft} style={{ marginTop: "5%" }}>
            <Checkbox label="Increment Revision ? " boxSide="end" onChange={this._onIncrementRevisionChecked} />
            </div>
        </div>
        <div className={styles.divrow}>
          <div style={{ width: "80%" }} >< TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentschange} multiline required autoAdjustHeight></TextField></div>
          <div><IconButton iconProps={AddIcon} title="Addindex" ariaLabel="Addindex" onClick={this._addindex} style={{ padding: "58px 0px 0px 45px" }} /></div>
        </div>
        <table className={styles.tableModal}  hidden={this.state.showGrid} >
                                <tr style={{background: "#f4f4f4"}}>
                                  <th style={{ padding: "5px 10px" }} >Slno</th>
                                  <th style={{ padding: "5px 10px" }}>DocumentIndex</th>
                                  <th style={{ padding: "5px 10px" }}>SubContractor Doc No</th>
                                 <th style={{ padding: "5px 10px" }}>ReceivedDate</th>
                                  <th style={{ padding: "5px 10px" }}>Owner</th>
                                  <th style={{ padding: "5px 10px" }}>Comments</th>             
                                  <th style={{ padding: "5px 10px" }}>Delete</th>
                                </tr>
                                {this.state.gridDocument.map((items,key)=>{
                                  return(
                                    <tr style={{borderBottom:"1px solid #f4f4f4"}}>
                                    <td style={{ padding: "5px 10px" }}>{key+1}</td>
                                     <td style={{ padding: "5px 10px" }}>{items.DocumentIndex} </td>
                                    <td style={{ padding: "5px 10px" }}>{items.SubContractorDocId} </td>
                                    <td style={{ padding: "5px 10px" }}>{items.ReceivedDate}</td>
                                    <td style={{ padding: "5px 10px" }}>{items.Owner} </td>
                                    <td style={{ padding: "5px 10px" }}>{items.Comments}</td>       
                                    <td style={{ padding: "5px 10px" }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={()=>this.itemDeleteFromGrid(items,key)}/></td>
                                  </tr>
                                  );
                                })}
                               
                          </table>
        <hr />
        <Label>Additional Document </Label>
        <div className={styles.divrow}>
          <div className={styles.wdthrgt}>
            <Label >Upload Document:</Label>
            <input type="file" name="externalFile" id="externalFile"></input>
          </div>
          <div className={styles.wdthlft}><DatePicker label="Date" value={this.state.externalDate} onSelectDate={this._onDatePickerChange} placeholder="Select a date" /></div>
         </div>
         <div className={styles.divrow}>
          <div style={{ width: "80%" }} >< TextField label="Comments" id="comments" value={this.state.externalComments} onChange={this._externalCommentsChange} multiline required autoAdjustHeight></TextField></div>
          <div><IconButton iconProps={AddIcon} title="Addindex" ariaLabel="Addindex" onClick={this._addexternalindex} style={{ padding: "58px 0px 0px 45px" }} /></div>
        </div>
        <table className={styles.tableModal}  hidden={this.state.showExternalGrid} >
                                <tr style={{background: "#f4f4f4"}}>
                                  <th style={{ padding: "5px 10px" }} >Slno</th>
                                  <th style={{ padding: "5px 10px" }}>Doc Name</th>
                                  <th style={{ padding: "5px 10px" }}>ReceivedDate</th>
                                  <th style={{ padding: "5px 10px" }}>Comments</th>             
                                  <th style={{ padding: "5px 10px" }}>Delete</th>
                                </tr>
                                {this.state.gridExternalDocument.map((items,key)=>{
                                  return(
                                    <tr style={{borderBottom:"1px solid #f4f4f4"}}>
                                    <td style={{ padding: "5px 10px" }}>{key+1}</td>
                                     <td style={{ padding: "5px 10px" }}>{items.DocName} </td>
                                   <td style={{ padding: "5px 10px" }}>{items.ExternalDate}</td>
                                    <td style={{ padding: "5px 10px" }}>{items.Comments}</td>       
                                    <td style={{ padding: "5px 10px" }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={()=>this.itemDeleteFromGrid(items,key)}/></td>
                                  </tr>
                                  );
                                })}
                               
                          </table>
        <DialogFooter>


                <div className={styles.rgtalign}>
                  <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
                </div>
                <div className={styles.rgtalign} >
                  <DefaultButton id="b2" className={styles.btn} onClick={this._saveAsDraft}>Save as draft</DefaultButton >

                  <DefaultButton id="b2" className={styles.btn} onClick={this._submit}>Submit</DefaultButton >
                  <DefaultButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</DefaultButton >
                </div>
              </DialogFooter>
      </div>
    );
  }
}
