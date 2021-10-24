import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmecInboundSubContractorWebPartStrings';
import EmecInboundSubContractor from './components/EmecInboundSubContractor';
import { IEmecInboundSubContractorProps } from './components/IEmecInboundSubContractorProps';
import { sp } from '@pnp/sp';


export default class EmecInboundSubContractorWebPart extends BaseClientSideWebPart<IEmecInboundSubContractorProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IEmecInboundSubContractorProps> = React.createElement(
      EmecInboundSubContractor,
      {
        context: this.context,
        description: this.properties.description,
        redirectUrl:this.properties.redirectUrl,
        siteUrl:this.context.pageContext.web.serverRelativeUrl,
        projectInformationListName:this.properties.projectInformationListName,
        revisionLevelList:this.properties.revisionLevelList,
        transmittalCodeSettings:this.properties.transmittalCodeSettings,
        hubUrl:this.properties.hubUrl,
        hubsite:this.properties.hubsite,
        companyList:this.properties.companyList,
        transmittalOutlookLibrary:this.properties.transmittalOutlookLibrary,
        documentIndexList:this.properties.documentIndexList
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
                PropertyPaneTextField('hubUrl', {
                  label: 'Hub Url'
                }),
                PropertyPaneTextField('hubsite', {
                  label: 'Hubsite'
                }),
                PropertyPaneTextField('companyList', {
                  label: 'Company List'
                })
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('projectInformationListName', {
                  label: 'Project Information'
                }),
                PropertyPaneTextField('revisionLevelList', {
                  label: 'Revision Level List'
                }),
                PropertyPaneTextField('transmittalCodeSettings', {
                  label: 'Transmittal Code Settings'
                }),
                PropertyPaneTextField('transmittalOutlookLibrary', {
                  label: 'Transmittal Outlook Library'
                }),
                PropertyPaneTextField('documentIndexList', {
                  label: 'Document Index List'
                })
              ]
            }
           
          ]
        }
      ]
    };
  }
}
