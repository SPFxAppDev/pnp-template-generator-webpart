import { Guid } from "@microsoft/sp-core-library";
import { IContentType, IField } from "./interfaces";
import { AttributeMapper, ElementBase } from "./ElementBase";
import { isNullOrEmpty, isset } from "@spfxappdev/utility";

export interface ContentType extends IContentType {}

export class ContentType extends ElementBase {

    private readonly UniqeId: string;

    public elementName: string = 'pnp:ContentType';

    public FieldRefIds: string[] = [];

    private allSiteColumns: IField[];

    private get overwriteCT(): boolean {
        //TODO: Add logic to let the user decide
        return true;
    }

    public get ID(): string {
        if(isNullOrEmpty(this.ParentID)) {
            return '';
        }

        return this.ParentID + '00' + this.UniqeId;
    }

    protected attributeMapper: AttributeMapper[] = [
        { ownPropertyName: 'ID', elementName: 'ID' },
        { ownPropertyName: 'Name', elementName: 'Name' },
        { ownPropertyName: 'Description', elementName: 'Description' },
        { ownPropertyName: 'Group', elementName: 'Group' },
        { ownPropertyName: 'overwriteCT', elementName: 'Overwrite' }
    ];

    constructor() {
        super();

        this.ParentID = '0x01';
        this.UniqeId = Guid.newGuid().toString().toUpperCase().ReplaceAll('-', '');
    }

    public toElement(rootDocument: XMLDocument): Element {
        const element = super.toElement(rootDocument);
        
        if (!isNullOrEmpty(this.FieldRefIds)) {
            const fields = rootDocument.createElement("pnp:FieldRefs");
            this.FieldRefIds.forEach((fieldRefId: string) => {
                const sourceField = this.allSiteColumns.FirstOrDefault(f => f.ID == fieldRefId);

                if(isNullOrEmpty(sourceField)) {
                    return;
                }

                const field = rootDocument.createElement("pnp:FieldRef");
                field.setAttribute('ID', sourceField.ID);
                field.setAttribute('Required', sourceField.Required ? 'true' : 'false');
                field.setAttribute('Name', sourceField.Name);
                fields.appendChild(field);
            });
            
            element.appendChild(fields);
        }
        
        return element;
    }

    public getInternalIdentifier(): string {
        return this.UniqeId;
    }

    public setAllAvailableCustomSiteColumns(fields: IField[]): void {
        this.allSiteColumns = fields;
    }


}