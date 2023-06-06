import { IContentType, IField, IList } from "./interfaces";
import { AttributeMapper, ElementBase } from "./ElementBase";
import { isNullOrEmpty, isset } from "@spfxappdev/utility";

export interface List extends IList {}

export class List extends ElementBase {

    public elementName: string = 'pnp:ListInstance';

    public ContentTypeRefIds: string[] = [];

    private allSiteColumns: IField[];

    private allContentTypes: IContentType[];

    Title: string;
    Url: string;
    Description?: string;
    TemplateType: string;
    Hidden?: boolean;
    EnableAttachments?: boolean;
    EnableFolderCreation?: boolean;
    EnableVersioning?: boolean;
    MinorVersionLimit?: number;
    MaxVersionLimit?: number;
    DraftVersionVisibility?: number;

    protected attributeMapper: AttributeMapper[] = [
        { ownPropertyName: 'Title', elementName: 'Title' },
        { ownPropertyName: 'Url', elementName: 'Url' },
        { ownPropertyName: 'Description', elementName: 'Description' },
        { ownPropertyName: 'TemplateType', elementName: 'TemplateType' },
        { ownPropertyName: 'Hidden', elementName: 'Hidden' },
    ];

    constructor() {
        super();
        this.TemplateType = "100";
    }

    public toElement(rootDocument: XMLDocument): Element {
        const element = super.toElement(rootDocument);
        
        if (!isNullOrEmpty(this.ContentTypeRefIds)) {
            const contentTypeBindings = rootDocument.createElement("pnp:ContentTypeBindings");

            this.ContentTypeRefIds.forEach((ctRefId: string) => {
                const sourceContentType = this.allContentTypes.FirstOrDefault(f => f.ID == ctRefId);

                if(isNullOrEmpty(sourceContentType)) {
                    return;
                }

                const contentTypeBinding = rootDocument.createElement("pnp:ContentTypeBinding");
                contentTypeBinding.setAttribute('ContentTypeID', sourceContentType.ID);
                contentTypeBindings.appendChild(contentTypeBinding);
            });

            element.appendChild(contentTypeBindings);

        }
        
        return element;
    }

    public setAvailableFieldsAndContentTypes(fields: IField[], contentTypes: IContentType[]): void {
        this.allSiteColumns = fields;
        this.allContentTypes = contentTypes;
    }
}