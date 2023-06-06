import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PnPTemplateGeneratorWebPartStrings';
import PnPTemplateGenerator from './components/PnPTemplateGenerator';
import { IPnPTemplateGeneratorProps } from './components/IPnPTemplateGeneratorProps';
import { PnPTemplateGeneratorServiceService, IPnPTemplateGeneratorServiceService } from '../../services/PnPTemplateGeneratorService';

export interface IPnPTemplateGeneratorWebPartProps {
  description: string;
}

export default class PnPTemplateGeneratorWebPart extends BaseClientSideWebPart<IPnPTemplateGeneratorWebPartProps> {

  private pnpTemplateService: IPnPTemplateGeneratorServiceService;

  public render(): void {
    const element: React.ReactElement<IPnPTemplateGeneratorProps> = React.createElement(
      PnPTemplateGenerator,
      {
        pnpTemplateGeneratorService: this.pnpTemplateService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    this.pnpTemplateService = this.context.serviceScope.consume(PnPTemplateGeneratorServiceService.serviceKey);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
