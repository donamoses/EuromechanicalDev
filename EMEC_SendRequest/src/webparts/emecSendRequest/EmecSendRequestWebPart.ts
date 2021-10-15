import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmecSendRequestWebPartStrings';
import EmecSendRequest from './components/EmecSendRequest';
import { IEmecSendRequestProps } from './components/IEmecSendRequestProps';
import { sp } from '@pnp/sp';

export interface IEmecSendRequestWebPartProps {
  description: string;
  RedirectUrl: string;
}

export default class EmecSendRequestWebPart extends BaseClientSideWebPart<IEmecSendRequestProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IEmecSendRequestProps> = React.createElement(
      EmecSendRequest,
      {
        context: this.context,
        siteUrl:this.context.pageContext.web.serverRelativeUrl,
        hubUrl:this.properties.hubUrl,
        RedirectUrl:this.properties.RedirectUrl,
        project: this.properties.project,
        notificationPreference:this.properties.notificationPreference,
        emailNotification:this.properties.emailNotification,
        userMessageSettings:this.properties.userMessageSettings,
        WorkflowHeaderList:this.properties.WorkflowHeaderList,
        DocumentIndexList:this.properties.DocumentIndexList,
        WorkflowDetailsList:this.properties.WorkflowDetailsList,
        SourceDocumentLibrary:this.properties.SourceDocumentLibrary,
        DocumentRevisionLogList:this.properties.DocumentRevisionLogList,
        TransmittalCodeSettingsList:this.properties.TransmittalCodeSettingsList,
        WorkflowTasksList:this.properties.WorkflowTasksList,
        RevisionLevelList:this.properties.RevisionLevelList,
        TaskDelegationSettings:this.properties.TaskDelegationSettings,
        RevisionHistoryPage:this.properties.RevisionHistoryPage,
        DocumentApprovalPage:this.properties.DocumentApprovalPage,
        DocumentReviewPage:this.properties.DocumentReviewPage,
        AccessGroups:this.properties.AccessGroups,
        DepartmentList:this.properties.DepartmentList
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
               PropertyPaneTextField('DocumentIndexList',{
                  label:'Document Index List'
                }),
                PropertyPaneTextField('SourceDocumentLibrary',{
                  label: 'Source Document Library'
                }),
                PropertyPaneTextField('WorkflowHeaderList',{
                  label:'WorkflowHeaderList'
                }),
                PropertyPaneTextField('WorkflowDetailsList',{
                  label:'Workflow Details List'
                }),
               PropertyPaneTextField('DocumentRevisionLogList',{
                  label:'Document RevisionLog List'
                }),
                PropertyPaneTextField('RedirectUrl', {
                  label: 'Redirect Url'
                }),
              ]
            },
            {
              groupName: "HubSite",
              groupFields: [
                PropertyPaneTextField('hubUrl',{
                  label:'HubUrl'
                }),
                PropertyPaneTextField('AccessGroups',{
                  label:'Access Groups List'
                }),
                PropertyPaneTextField('notificationPreference',{
                  label:'Notification Preference'
                }),
                PropertyPaneTextField('emailNotification',{
                  label:'Email Notification'
                }),
                PropertyPaneTextField('userMessageSettings',{
                  label:'User Message Settings'
                }),
                PropertyPaneTextField('WorkflowTasksList',{
                  label:'Workflow Tasks List'
                }),
                PropertyPaneTextField('TaskDelegationSettings',{
                  label:'Task Delegation Settings'
                }),
                PropertyPaneTextField('TaskDelegationSettings',{
                  label:'Task Delegation Settings'
                }),
              ]
            },
            {
              groupName: "Pages",
              groupFields: [
                PropertyPaneTextField('DocumentReviewPage',{
                  label:'Document Review Page'
                }),
                PropertyPaneTextField('DocumentApprovalPage',{
                  label:'Document Approval Page'
                }),
                PropertyPaneTextField('RevisionHistoryPage',{
                  label:'Revision History Page'
                }),
              ]
            },
    
            {
              groupName: "Project",
              groupFields: [
                PropertyPaneToggle('project',{
                  label:'Project',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('RevisionLevelList',{
                  label:'Revision Level List'
                }),
                PropertyPaneTextField('DepartmentList',{
                  label:'Department List'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
