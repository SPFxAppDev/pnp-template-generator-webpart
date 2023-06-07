import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
// import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
// import { PageContext } from '@microsoft/sp-page-context';
import { isNullOrEmpty } from '@spfxappdev/utility';
import { PnPTemplate } from '../models/PnPTemplate';
import { IField, IContentType, IList } from '../models';



export interface IPnPTemplateGeneratorServiceService {
    getTemplate(): string;
    // addList(): void;
    // addSiteColumn(): void;
    pnpTemplate: PnPTemplate;
}

type TaxonomyElementArrayProps = { NameElementValue: string; ValueElement: { value: string, attributes: Record<string, string> } };


export class PnPTemplateGeneratorServiceService implements IPnPTemplateGeneratorServiceService {

    public static readonly serviceKey: ServiceKey<IPnPTemplateGeneratorServiceService> =
        ServiceKey.create<PnPTemplateGeneratorServiceService>('SPFxAppDev:IPnPTemplateGeneratorServiceService', PnPTemplateGeneratorServiceService);

    // private spHttpClient: SPHttpClient;
    // private pageContext: PageContext;

    private tpl: XMLDocument;

    private provisioningTemplate: Element;

    public pnpTemplate: PnPTemplate;

    // private lists: any[] = [];

    constructor(serviceScope: ServiceScope) {
        this.pnpTemplate = new PnPTemplate();

        // serviceScope.whenFinished(() => {
        //     this.spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
        //     this.pageContext = serviceScope.consume(PageContext.serviceKey);
        // });
    }

    public getTemplate(): string {
        this.generateTemplate();


        const sourceXml = this.tpl.documentElement.outerHTML;
        const xmlDoc = new DOMParser().parseFromString(sourceXml, 'application/xml');
        const xsltDoc = new DOMParser().parseFromString([
            // describes how we want to modify the XML - indent everything
            '<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform">',
            '  <xsl:strip-space elements="*"/>',
            '  <xsl:template match="para[content-style][not(text())]">', // change to just text() to strip space in text nodes
            '    <xsl:value-of select="normalize-space(.)"/>',
            '  </xsl:template>',
            '  <xsl:template match="node()|@*">',
            '    <xsl:copy><xsl:apply-templates select="node()|@*"/></xsl:copy>',
            '  </xsl:template>',
            '  <xsl:output indent="yes"/>',
            '</xsl:stylesheet>',
        ].join('\n'), 'application/xml');

        const xsltProcessor = new XSLTProcessor();    
        xsltProcessor.importStylesheet(xsltDoc);
        const resultDoc = xsltProcessor.transformToDocument(xmlDoc);
        const resultXml = new XMLSerializer().serializeToString(resultDoc);
        return resultXml;

            // return this.tpl.documentElement.outerHTML;
    }

    private addLists(): void {
        const parentSelector = 'pnp:Lists';
        const commentStart = this.tpl.createComment(" Lists ~ ");
        this.provisioningTemplate.appendChild(commentStart);

        this.createTemplateElementCollection(parentSelector, this.provisioningTemplate);

        const lists = this.provisioningTemplate.getElementsByTagName(parentSelector)[0];

        // const list = this.tpl.createElement("pnp:ListInstance");
        // this.createTemplateElementCollection('pnp:ContentTypeBindings', list);

        this.pnpTemplate.lists.forEach((list: IList) => {
            list.setAvailableFieldsAndContentTypes(this.pnpTemplate.siteColumns, this.pnpTemplate.contentTypes);
            const listElement = list.toElement(this.tpl);
            lists.appendChild(listElement);
        });

        // lists.appendChild(list);

        const commentEnd = this.tpl.createComment(" ~ Lists ");
        this.provisioningTemplate.appendChild(commentEnd);
    }

    private addSiteColumns(): void {
        const parentSelector = 'pnp:SiteFields';
        const commentStart = this.tpl.createComment(" Site Columns ~ ");
        this.provisioningTemplate.appendChild(commentStart);
        this.createTemplateElementCollection(parentSelector, this.provisioningTemplate);

        const fields = this.provisioningTemplate.getElementsByTagName(parentSelector)[0];

        this.pnpTemplate.siteColumns.forEach((field: IField) => {

            const fieldElement = field.toElement(this.tpl);

            if(field.Type === 'TaxonomyFieldType') {
                this.createAdditionalElementsForTaxonomy(field, this.tpl, fieldElement);
                
            }

            const fieldComment = this.tpl.createComment(` ${field.Name} `);
            fields.appendChild(fieldComment);
            fields.appendChild(fieldElement);
        });

        const commentEnd = this.tpl.createComment(" ~ Site Columns ");
        this.provisioningTemplate.appendChild(commentEnd);
    }

    private addContentTypes(): void {
        const parentSelector = 'pnp:ContentTypes';
        const commentStart = this.tpl.createComment(" Site Content Types ~ ");
        this.provisioningTemplate.appendChild(commentStart);
        this.createTemplateElementCollection(parentSelector, this.provisioningTemplate);

        const contentTypes = this.provisioningTemplate.getElementsByTagName(parentSelector)[0];
        
        this.pnpTemplate.contentTypes.forEach((contentType: IContentType) => {

            contentType.setAllAvailableCustomSiteColumns(this.pnpTemplate.siteColumns);
            const contentTypeElement = contentType.toElement(this.tpl);
            contentTypes.appendChild(contentTypeElement);
        });

        const commentEnd = this.tpl.createComment(" ~ Site Content Types ");
        this.provisioningTemplate.appendChild(commentEnd);
    }

    private generateTemplate(): void {
        this.createEmptyTemplateXml();

        if(!isNullOrEmpty(this.pnpTemplate.siteColumns)) {
            this.addSiteColumns();
        }

        if(!isNullOrEmpty(this.pnpTemplate.contentTypes)) {
            this.addContentTypes();
        }

        if(!isNullOrEmpty(this.pnpTemplate.lists)) {
            this.addLists();
        }
    }

    private createEmptyTemplateXml(): void {
        this.tpl = document.implementation.createDocument(
            "http://schemas.dev.office.com/PnP/2022/09/ProvisioningSchema",
            "pnp:Provisioning",
            null
        );

        this.tpl.documentElement.setAttribute('Author', 'SPFx-App.dev');
        this.tpl.documentElement.setAttribute('Generator', 'SPFx-App.dev');
        this.tpl.documentElement.setAttribute('Version', '1.0');

        const preferences = this.tpl.createElement("pnp:Preferences");
        preferences.setAttribute('Author', 'SPFx-App.dev');
        preferences.setAttribute('Generator', 'SPFx-App.dev');
        preferences.setAttribute('Version', '1.0');

        const templates = this.tpl.createElement("pnp:Templates");
        // templates.setAttribute("Generator", "PnP.Framework, Version=1.8.3.0, Culture=neutral, PublicKeyToken=0d501f89f11b748c");
        templates.setAttribute('ID', 'SPFxAppDev-Template-Generator');

        this.provisioningTemplate = this.tpl.createElement("pnp:ProvisioningTemplate");
        this.provisioningTemplate.setAttribute('ID', 'SPFxAppDev-Template-Generator');

        templates.appendChild(this.provisioningTemplate);

        this.tpl.documentElement.appendChild(preferences);
        this.tpl.documentElement.appendChild(templates);

    }

    private createTemplateElementCollection(elementName: string, parentElement: Element): void {

        const elements = parentElement.getElementsByTagName(elementName);

        if(isNullOrEmpty(elements) || elements.length > 0) {
            return;
        }

        const el = this.tpl.createElement(elementName);
        parentElement.appendChild(el);
    }

    private createAdditionalElementsForTaxonomy(taxonomyField: IField, rootDocument: XMLDocument, taxonomyElement: Element): void {

        
        const customization = rootDocument.createElement("Customization");
        const propsArray = rootDocument.createElement("ArrayOfProperty");

        const props: TaxonomyElementArrayProps[] = [
            { 
                NameElementValue: 'SspId', 
                ValueElement: { 
                    value: '{sitecollectiontermstoreid}', 
                    attributes: { 
                        'xmlns:q1': 'http://www.w3.org/2001/XMLSchema',
                        'p4:type': 'q1:string',
                        'xmlns:p4': 'http://www.w3.org/2001/XMLSchema-instance',
                    }
                } 
            },
            { 
                NameElementValue: 'UserCreated', 
                ValueElement: { 
                    value: 'false', 
                    attributes: { 
                        'xmlns:q4': 'http://www.w3.org/2001/XMLSchema',
                        'p4:type': 'q4:boolean',
                        'xmlns:p4': 'http://www.w3.org/2001/XMLSchema-instance',
                    }
                } 
            },
            { 
                NameElementValue: 'Open', 
                ValueElement: { 
                    value: 'false', 
                    attributes: { 
                        'xmlns:q5': 'http://www.w3.org/2001/XMLSchema',
                        'p4:type': 'q5:boolean',
                        'xmlns:p4': 'http://www.w3.org/2001/XMLSchema-instance',
                    }
                } 
            },
            { 
                NameElementValue: 'IsPathRendered', 
                ValueElement: { 
                    value: 'false', 
                    attributes: { 
                        'xmlns:q7': 'http://www.w3.org/2001/XMLSchema',
                        'p4:type': 'q7:boolean',
                        'xmlns:p4': 'http://www.w3.org/2001/XMLSchema-instance',
                    }
                } 
            },
            { 
                NameElementValue: 'IsKeyword', 
                ValueElement: { 
                    value: 'false', 
                    attributes: { 
                        'xmlns:q8': 'http://www.w3.org/2001/XMLSchema',
                        'p4:type': 'q8:boolean',
                        'xmlns:p4': 'http://www.w3.org/2001/XMLSchema-instance',
                    }
                } 
            },
            { 
                NameElementValue: 'CreateValuesInEditForm', 
                ValueElement: { 
                    value: 'false', 
                    attributes: { 
                        'xmlns:q10': 'http://www.w3.org/2001/XMLSchema',
                        'p4:type': 'q10:boolean',
                        'xmlns:p4': 'http://www.w3.org/2001/XMLSchema-instance',
                    }
                } 
            }
        ];

        if(!isNullOrEmpty(taxonomyField.TermGroupName) && !isNullOrEmpty(taxonomyField.TermSetName)) {
            props.AddAt(1, { 
                NameElementValue: 'TermSetId', 
                ValueElement: { 
                    value: `{termsetid:${taxonomyField.TermGroupName}:${taxonomyField.TermSetName}}`, 
                    attributes: { 
                        'xmlns:q2': 'http://www.w3.org/2001/XMLSchema',
                        'p4:type': 'q2:string',
                        'xmlns:p4': 'http://www.w3.org/2001/XMLSchema-instance',
                    }
                } 
            })
        }

        props.forEach((props: TaxonomyElementArrayProps) => {
            const propertyElement = rootDocument.createElement("Property");

            const nameElement = rootDocument.createElement("Name");
            nameElement.innerHTML = props.NameElementValue;

            const valueElement = rootDocument.createElement("Value");
            valueElement.innerHTML = props.ValueElement.value;

            Object.getOwnPropertyNames(props.ValueElement.attributes).forEach((attributeName: string) => {
                valueElement.setAttribute(attributeName, props.ValueElement.attributes[attributeName]);
            });

            propertyElement.appendChild(nameElement);
            propertyElement.appendChild(valueElement);

            propsArray.appendChild(propertyElement);
        })


        customization.appendChild(propsArray);
        taxonomyElement.appendChild(customization);
    }

}