import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmecApprovalWpWebPartStrings';
import EmecApprovalWp from './components/EmecApprovalWp';
import { IEmecApprovalWpProps } from './components/IEmecApprovalWpProps';
import { sp } from '@pnp/sp';

export default class EmecApprovalWpWebPart extends BaseClientSideWebPart<IEmecApprovalWpProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IEmecApprovalWpProps> = React.createElement(
      EmecApprovalWp,
      {
        context: this.context,
        description: this.properties.description,
        project:this.properties.project,
        RedirectUrl:this.properties.RedirectUrl,
        siteUrl:this.context.pageContext.web.serverRelativeUrl,
        hubUrl:this.properties.hubUrl,
        notificationPreference:this.properties.notificationPreference,
        emailNotification:this.properties.emailNotification,
        userMessageSettings:this.properties.userMessageSettings,
        WorkflowHeaderList:this.properties.WorkflowHeaderList,
        DocumentIndexList:this.properties.DocumentIndexList,
        WorkflowDetailsList:this.properties.WorkflowDetailsList,
        SourceDocument:this.properties.SourceDocument,
        PublishedDocument:this.properties.PublishedDocument,
        DocumentRevisionLogList:this.properties.DocumentRevisionLogList,
        TransmittalCodeSettingsList:this.properties.TransmittalCodeSettingsList,
        WorkflowTasksList:this.properties.WorkflowTasksList,
        AccessGroups:this.properties.AccessGroups,
        DepartmentList:this.properties.DepartmentList,
        SourceDocumentLibrary:this.properties.SourceDocumentLibrary
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
              groupName: "Hub Site",
              groupFields: [
                PropertyPaneTextField('hubUrl',{
                  label:'HubUrl'
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
                PropertyPaneTextField('AccessGroups',{
                  label:'Access Groups List'
                }),
                PropertyPaneTextField('WorkflowTasksList',{
                  label:'Workflow Tasks List'
                }),
                PropertyPaneTextField('DepartmentList',{
                  label:'Department List'
                }),
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('DocumentRevisionLogList',{
                  label:'Document RevisionLog List'
                }),
                PropertyPaneTextField('DocumentIndexList',{
                  label:'Document Index List'
                }),
                PropertyPaneTextField('SourceDocument',{
                  label:'Source Document Library'
                }),
                PropertyPaneTextField('SourceDocumentLibrary',{
                  label:'Source Document View Library'
                }),
                PropertyPaneTextField('WorkflowHeaderList',{
                  label:'WorkflowHeaderList'
                }),
                PropertyPaneTextField('WorkflowDetailsList',{
                  label:'Workflow Details List'
                }),
               PropertyPaneTextField('PublishedDocument',{
                  label:'Published Document Library'
                }),
                PropertyPaneTextField('RedirectUrl', {
                  label: 'Redirect Url'
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
                PropertyPaneTextField('TransmittalCodeSettingsList',{
                  label:'Transmittal Code Settings List'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
