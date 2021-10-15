import * as React from 'react';
import styles from './RevisionCoding.module.scss';
import { IRevisionCodingProps } from './IRevisionCodingProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, TextField } from 'office-ui-fabric-react';
import { sp } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import replaceString from 'replace-string';
export interface IRevisionState {
  revision: any;

}
export default class RevisionCoding extends React.Component<IRevisionCodingProps, IRevisionState, {}> {
  public constructor(props: IRevisionCodingProps) {
    super(props);
    this.state = {
      revision: ""

    };

  }
  public _revisioncoding = async () => {
    console.log(this.state.revision);
    let RevisionItemId;
    let currentrev;
    var xfirstletter;
    var Prefix;
    var MulPrefix = "";
    var Prefixyes;
    let revisionnow;
    var revs;
    let revision = this.state.revision;
    let increment;
    let init;
    let result;
    let next;
    let ptype;
    const HeaderItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/WorkflowHeader").items.getById(1).select("RevisionCodingId,ApproveInSameRevision").get();
    RevisionItemId = HeaderItem.RevisionCodingId;
    const RevisionItems: any = await sp.web.getList(this.props.siteUrl + "/Lists/RevisionSettings").items.getById(RevisionItemId).get();
    console.log(RevisionItems);
    if (RevisionItems.AutoIncrement == "TRUE") {
      if(RevisionItems.Pattern != null){
      let Incrementarray: any[] = RevisionItems.Pattern.split('+');
      ptype = Incrementarray[0];
     
      increment = parseInt(Incrementarray[1]);
      }
      if (RevisionItems.Types == "Numeric") {
        if (RevisionItems.MinN != null) {
          if (parseInt(this.state.revision) >= parseInt(RevisionItems.MinN)) {
            currentrev = parseInt(revision) + increment;
            revisionnow = currentrev.toString();
            alert(revisionnow);
          }
          else {
            currentrev = RevisionItems.MinN;
            revisionnow = currentrev.toString();
            alert(revisionnow);
          }
        }
      }
      if (RevisionItems.Types == "Alpha") {
        let StartWith = RevisionItems.StartWith.charCodeAt(0);
        let charcode = revision.charCodeAt(0);
        if (charcode >= StartWith) {
          if (charcode == 90 || charcode == 122) {
            alert("Please select valid pattern");
          }
          else {
            revisionnow = String.fromCharCode(charcode + increment);
            alert(revisionnow);
          }
        }
        else {
          revisionnow = RevisionItems.StartWith;
          alert(revisionnow);
        }
      }
      if (RevisionItems.Types == "Prefix") {
        let revlength = parseInt(revision.length);
      
       
        for (let i = 0; i < revlength; i++) {
          Prefix = MulPrefix + revision.charAt(i);
          let prefixlength = parseInt(Prefix.length);
          if (Prefix == RevisionItems.StartPrefix) {
            xfirstletter = revision.slice(prefixlength);
            Prefixyes = "Yes";
          }
          else {
            MulPrefix = Prefix;
          }
        }
        if (parseInt(this.state.revision) >= parseInt(RevisionItems.MinN)) {
          currentrev = parseInt(revision) + increment;
         
          if(currentrev <10){
            revisionnow = "0"+currentrev.toString();
          }
          else{
            revisionnow = currentrev.toString();
          }
          alert(revisionnow);
        }
        else if (Prefixyes == "Yes" && RevisionItems.MinN != null) {
          if (parseInt(xfirstletter) >= parseInt(RevisionItems.MinN)) {
            let intxfirstletter = parseInt(xfirstletter);
            currentrev = intxfirstletter + increment;
            revs = currentrev.toString();
            revisionnow = RevisionItems.StartPrefix + revs;
            alert(revisionnow);
          }
          else {
            revisionnow = RevisionItems.StartPrefix + RevisionItems.MinN;
            alert(revisionnow);
          }
        }
        else {
          revisionnow = RevisionItems.StartPrefix + RevisionItems.MinN;
          alert(revisionnow);
        }
      
      }
      if (RevisionItems.Types == "MultiPrefix") {
        let startArray: any[] = RevisionItems.StartPrefix.split(',');
        let revlength = parseInt(revision.length);
        for (let i = 0; i < revlength; i++) {
          Prefix = MulPrefix + revision.charAt(i);
          result = startArray.indexOf(Prefix);
          let start = startArray[result];
          let prefixlength = parseInt(Prefix.length);
          if (Prefix == start) {
            xfirstletter = revision.slice(prefixlength);
            Prefixyes = "Yes";
            init = Prefix;
            next = result + 1;
          }
          else {
            MulPrefix = Prefix;
          }
        }

        if (parseInt(this.state.revision) >= parseInt(RevisionItems.MinN)&& parseInt(this.state.revision) < parseInt(RevisionItems.MaxN)) {
          
          currentrev = parseInt(revision) + increment;
          revisionnow = currentrev.toString();
          alert(revisionnow);
        }
        else if(parseInt(this.state.revision) == parseInt(RevisionItems.MaxN)){
          result = startArray.indexOf(" ");
          next = result + 1;
          revisionnow = startArray[next] + RevisionItems.MinN;
          alert(revisionnow);
        }
        else if (Prefixyes == "Yes" && RevisionItems.MinN != null) {
          if (parseInt(xfirstletter) >= parseInt(RevisionItems.MinN) && parseInt(xfirstletter) < parseInt(RevisionItems.MaxN)) {
            let intxfirstletter = parseInt(xfirstletter);
            currentrev = intxfirstletter + increment;
            revs = currentrev.toString();
            revisionnow = init + revs;
            alert(revisionnow);
          }
          else if(startArray.length == next){
            alert("Please select valid pattern");
          }
          else {
            revisionnow = startArray[next] + RevisionItems.MinN;
            alert(revisionnow);
          }
        }
        else {
          revisionnow = startArray[0] + RevisionItems.MinN;
          alert(revisionnow);
        }
      }
      if (RevisionItems.Types == "SingleStart") {
        let revlength = parseInt(revision.length);
        let charcode = revision.charCodeAt(0);
        if(revision == RevisionItems.StartWith){
          if(RevisionItems.StartPrefix != null && RevisionItems.MinN != null){
            revisionnow = RevisionItems.StartPrefix+RevisionItems.MinN;
            alert(revisionnow);
          }
          else if(RevisionItems.StartPrefix == null && RevisionItems.MinN != null){
            revisionnow = RevisionItems.MinN;
            alert(revisionnow);
          }
        }
        else if(RevisionItems.StartPrefix != null && ptype == "N" ){
        for (let i = 0; i < revlength; i++) {
          Prefix = MulPrefix + revision.charAt(i);
          let prefixlength = parseInt(Prefix.length);
          if (Prefix == RevisionItems.StartPrefix) {
            xfirstletter = revision.slice(prefixlength);
            Prefixyes = "Yes";
          }
          else {
            MulPrefix = Prefix;
          }
        }
         if (parseInt(this.state.revision) >= parseInt(RevisionItems.MinN)) {
          currentrev = parseInt(revision) + increment;
         
          if(currentrev <10){
            revisionnow = "0"+currentrev.toString();
          }
          else{
            revisionnow = currentrev.toString();
          }
          alert(revisionnow);
        }
        else if (Prefixyes == "Yes" && RevisionItems.MinN != null) {
          if (parseInt(xfirstletter) >= parseInt(RevisionItems.MinN)) {
            let intxfirstletter = parseInt(xfirstletter);
            currentrev = intxfirstletter + increment;
            revs = currentrev.toString();
            revisionnow = RevisionItems.StartPrefix + revs;
            alert(revisionnow);
          }
          else {
            revisionnow = RevisionItems.StartWith;
            alert(revisionnow);
          }
        }
        else {
          revisionnow = RevisionItems.StartWith;
          alert(revisionnow);
        }
        }
        else if(charcode == 90|| charcode ==122){
          alert("Please select valid pattern");
        }
        else if(ptype == "A" &&(charcode >= 65 && charcode <90)){
          revisionnow = String.fromCharCode(charcode + increment);
            alert(revisionnow);
        }
        else if(ptype == "N" && parseInt(RevisionItems.MinN) >= 0){
          if (RevisionItems.MinN != null) {
            if (parseInt(this.state.revision) >= parseInt(RevisionItems.MinN)) {
              currentrev = parseInt(revision) + increment;
              revisionnow = currentrev.toString();
              alert(revisionnow);
            }
            else if(RevisionItems.StartWith !=null){
               revisionnow = RevisionItems.StartWith;
            alert(revisionnow);
            }
            else if(RevisionItems.StartPrefix!=null){
            revisionnow = RevisionItems.StartPrefix+RevisionItems.MinN;
            alert(revisionnow);
            }     
            else {
              currentrev = RevisionItems.MinN;
              revisionnow = currentrev.toString();
              alert(revisionnow);
            }
          }
        }
        else{
          revisionnow = RevisionItems.StartWith;
            alert(revisionnow);
        }
   
      }
     if (RevisionItems.Types == "AlphaNumeric") {
        let StartWith = RevisionItems.StartWith.charCodeAt(0);
        let EndWith = RevisionItems.EndWith.charCodeAt(0);
        let charcode = revision.charCodeAt(0);
        if (charcode >= StartWith && charcode < EndWith) {
          revisionnow = String.fromCharCode(charcode + increment);
          alert(revisionnow);
        }
        else {
          if (RevisionItems.MinN != null) {
            if (parseInt(this.state.revision) >= parseInt(RevisionItems.MinN)) {
              currentrev = parseInt(revision) + increment;
              revisionnow = currentrev.toString();
              alert(revisionnow);
            }
            else {
              currentrev = RevisionItems.MinN;
              revisionnow = currentrev.toString();
              alert(revisionnow);
            }
          }
        }

      }
     if(RevisionItems.Types == "NumericAlpha"){
      let StartWith = RevisionItems.StartWith.charCodeAt(0);
      let charcode = revision.charCodeAt(0);
      let EndWith = RevisionItems.EndWith.charCodeAt(0);
      if (RevisionItems.MinN != null) {
        if (parseInt(this.state.revision) >= parseInt(RevisionItems.MinN)&&parseInt(this.state.revision) < parseInt(RevisionItems.MaxN) ) {
          currentrev = parseInt(revision) + increment;
          revisionnow = currentrev.toString();
          alert(revisionnow);
        }
        else if(parseInt(this.state.revision) >= parseInt(RevisionItems.MaxN)){
          
          revisionnow = RevisionItems.StartWith;
          alert(revisionnow);
        }
        else if(charcode >= StartWith && charcode < EndWith){
          revisionnow = String.fromCharCode(charcode + increment);
              alert(revisionnow);
        }
        
        else {
          currentrev = RevisionItems.MinN;
          revisionnow = currentrev.toString();
          alert(revisionnow);
        }
      }
     }
    }//AutoIncrement End
    else {
      let Pattern = RevisionItems.Pattern;
      let Patternarray: any[] = Pattern.split(',');
      let count = 0;
      
      for (let k = 0; k <= Patternarray.length; k++) {
        let rev = " "+revision;
        result = Patternarray.indexOf(rev);
        if (rev == Patternarray[k]|| revision == Patternarray[k]) {
          count = k + 1;
        }
        
      }
      if (count >= Patternarray.length||count == 0 ) {
        alert("Please Verify Pattern");
      }
      
      else {
        currentrev = Patternarray[count];
        revisionnow = currentrev;
        alert(revisionnow);
      }
    }
  }
  public revisionchange = async (ev: React.FormEvent<HTMLInputElement>, revision?: any) => {
    this.setState({
      revision: revision || '',

    });


  }
  public render(): React.ReactElement<IRevisionCodingProps> {
    return (
      <div className={styles.revisionCoding}>
        <TextField id="pin"
          onChange={this.revisionchange}
          placeholder="Revision"
          value={this.state.revision} ></TextField>
        <DefaultButton id="b1" onClick={this._revisioncoding}>Revision</DefaultButton >
      </div>
    );
  }
}
