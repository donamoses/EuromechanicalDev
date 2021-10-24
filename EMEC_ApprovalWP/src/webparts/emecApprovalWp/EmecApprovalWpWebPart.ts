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
        redirectUrl:this.properties.redirectUrl,
        siteUrl:this.context.pageContext.web.serverRelativeUrl,
        hubUrl:this.properties.hubUrl,
        notificationPreference:this.properties.notificationPreference,
        emailNotification:this.properties.emailNotification,
        userMessageSettings:this.properties.userMessageSettings,
        workflowHeaderList:this.properties.workflowHeaderList,
        documentIndexList:this.properties.documentIndexList,
        workflowDetailsList:this.properties.workflowDetailsList,
        sourceDocument:this.properties.sourceDocument,
        publishedDocument:this.properties.publishedDocument,
        documentRevisionLogList:this.properties.documentRevisionLogList,
        transmittalCodeSettingsList:this.properties.transmittalCodeSettingsList,
        workflowTasksList:this.properties.workflowTasksList,
        PermissionMatrixSettings:this.properties.PermissionMatrixSettings,
        departmentList:this.properties.departmentList,
        sourceDocumentLibrary:this.properties.sourceDocumentLibrary,
        siteAddress:this.properties.siteAddress,
        accessGroupDetailsList:this.properties.accessGroupDetailsList,
        hubsite:this.properties.hubsite,
        projectInformationListName:this.properties.projectInformationListName
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
                PropertyPaneTextField('hubsite',{
                  label:'hubsite'
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
                PropertyPaneTextField('PermissionMatrixSettings',{
                  label:'Permission Matrix Settings List'
                }),
                PropertyPaneTextField('workflowTasksList',{
                  label:'Workflow Tasks List'
                }),
                PropertyPaneTextField('departmentList',{
                  label:'Department List'
                }),
                PropertyPaneTextField('accessGroupDetailsList',{
                  label:'AccessGroupDetailsList'
                }),
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('documentRevisionLogList',{
                  label:'Document RevisionLog List'
                }),
                PropertyPaneTextField('documentIndexList',{
                  label:'Document Index List'
                }),
                PropertyPaneTextField('sourceDocument',{
                  label:'Source Document Library'
                }),
               
                PropertyPaneTextField('workflowHeaderList',{
                  label:'WorkflowHeaderList'
                }),
                PropertyPaneTextField('workflowDetailsList',{
                  label:'Workflow Details List'
                }),
               PropertyPaneTextField('publishedDocument',{
                  label:'Published Document Library'
                }),
                PropertyPaneTextField('redirectUrl', {
                  label: 'Redirect Url'
                }),
               ]
            },
            {
              groupName: "LA Params",
              groupFields: [
                PropertyPaneTextField('siteAddress',{
                  label:'SiteAddress'
                }),
                PropertyPaneTextField('sourceDocumentLibrary',{
                  label:'Source Document View Library'
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
                PropertyPaneTextField('transmittalCodeSettingsList',{
                  label:'Transmittal Code Settings List'
                }),
                PropertyPaneTextField('projectInformationListName',{
                  label:'projectInformationListName'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
