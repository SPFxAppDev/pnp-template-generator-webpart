import * as React from 'react';
import styles from './PnPTemplateGenerator.module.scss';
import { IPnPTemplateGeneratorProps } from './IPnPTemplateGeneratorProps';
import { MonacoEditor } from "@pnp/spfx-controls-react/lib/MonacoEditor";
import { DefaultButton, Label, Pivot, PivotItem, Spinner } from '@fluentui/react';
import SiteColumnsGenerator from './generator/SiteColumnsGenerator';
import ContentTypesGenerator from './generator/ContentTypesGenerator';
import ListGenerator from './generator/ListGenerator';

interface IPnPTemplateGeneratorState {
  showTemplateLoadingSpinner: boolean;
}

export default class PnPTemplateGenerator extends React.Component<IPnPTemplateGeneratorProps, IPnPTemplateGeneratorState> {

  public state: IPnPTemplateGeneratorState = {
    showTemplateLoadingSpinner: false
  };

  constructor(props: IPnPTemplateGeneratorProps) {
    super(props);
  }

  public render(): React.ReactElement<IPnPTemplateGeneratorProps> {   

    return (
      <div className={styles.pnpTemplateGenerator}>
        <Label className={styles['header-label']}>PnP Provisioning Template Generator</Label>
        <div>
        <Pivot aria-label="Basic Pivot Example">
          <PivotItem
            headerText="Site Columns"
          >
            <SiteColumnsGenerator 
              pnpTemplateGeneratorService={this.props.pnpTemplateGeneratorService} 
              onChange={() => {
                this.reloadTemplate();
              }}
            />
          </PivotItem>
          <PivotItem headerText="Content Types">
            <ContentTypesGenerator 
              pnpTemplateGeneratorService={this.props.pnpTemplateGeneratorService} 
              onChange={() => {
                this.reloadTemplate();
              }}
            />
          </PivotItem>
          <PivotItem headerText="Lists">
            <ListGenerator
              pnpTemplateGeneratorService={this.props.pnpTemplateGeneratorService} 
              onChange={() => {
                this.reloadTemplate();
              }}
            />
          </PivotItem>
        </Pivot>
        </div>
        <div>
          <Label className={styles['header-label']}>Generated PnP Template</Label>
          {this.state.showTemplateLoadingSpinner && <Spinner />}
          {!this.state.showTemplateLoadingSpinner && 
            <>
            <div className={styles['template-header']}>
              <DefaultButton 
                text='Copy to Clipboard'
                onClick={async () => {
                  try { await window.navigator.clipboard.writeText(this.props.pnpTemplateGeneratorService.getTemplate()) } 
                  catch {}
                  finally {}
                }} />

              <DefaultButton 
                text='Save template'
                onClick={() => {
                  this.saveTemplateAsXmlFile();
                }} />


            </div>
            <MonacoEditor value={this.props.pnpTemplateGeneratorService.getTemplate()}
               showMiniMap={true}
               readOnly={true}
               showLineNumbers={true}
               language={"xml"}/>
            </>
          }
        </div>
      </div>
    )
  }

  private reloadTemplate(): void {
    this.setState({
      showTemplateLoadingSpinner: true
    }, () => {
      this.setState({
        showTemplateLoadingSpinner: false
      });
    });
  }

  private saveTemplateAsXmlFile(): void {
    const link = document.createElement('a');
    link.download = 'template.xml';
    const blob = new Blob([this.props.pnpTemplateGeneratorService.getTemplate()], {type: 'text/xml'});
    link.href = window.URL.createObjectURL(blob);
    link.click();
  }
}
