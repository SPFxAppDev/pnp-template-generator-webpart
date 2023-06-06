import { Guid } from "@microsoft/sp-core-library";
import { IField } from "./interfaces";
import { AttributeMapper, ElementBase } from "./ElementBase";
import { isNullOrEmpty, isset } from "@spfxappdev/utility";

export interface Field extends IField {}

export class Field extends ElementBase {
    elementName: string = 'Field';

    ID: string;
    //Vorerst: Text, Boolean, Number, User(Multi), Lookup, TaxonomyFieldType(Multi), Note, DateTime, (Multi)Choice, URL, Currency?
    Type: string;
    Name: string;
    DisplayName: string;
    Description?: string;
    Format?: string;
    Required?: boolean = false;

    Group?: string;
    Hidden?: boolean = false;
    DefaultValue?: string;
    //TODO: Check which type can have INDEXED=TRUE Attribute
    Indexed?: boolean;

    //Choice
    Choices?: string[];
    FillInChoice?: boolean;

    //Taxonomy + UserMulti + Choice
    Mult?: boolean;

    //Note
    RichText?: boolean;
    get RichTextMode(): string {
        
        if(this.RichText) {
            return 'FullHtml'
        }

        return null;
    }

    NumLines?: number;

    //Lookup
    List?: string;
    ShowField?: string;

    //TaxonomyFieldTypeMulti
    public get SourceID(): string {
        if(this.Type === 'TaxonomyFieldType') {
            return 'http://schemas.microsoft.com/sharepoint/v3';
        }

        return null;
    }

    //User
    UserSelectionMode?: string;

    private allAllowedFormatsForTypes: Record<string, string[]> = {
        'Choice': ['Dropdown', 'RadioButtons'],
        'DateTime': ['DateOnly', 'DateTime'],
        'URL': ['Hyperlink', 'Image'],
        'Lookup': ['Dropdown']
    }

    // eslint-disable-line
    private get typeInXml(): string {

        if(!isset(this.multiAttribute)) {
            return this.Type;
        }

        if(!this.multiAttribute) {
            return this.Type;
        }

        if(this.Type === 'TaxonomyFieldType' || this.Type === 'User') {
            return this.Type + 'Multi';
        }

        if(this.Type === 'Choice') {
            return 'Multi' + this.Type;
        }

        return this.Type;
    }

    // eslint-disable-line
    private get multiAttribute(): boolean {
        
        if(this.Type === 'TaxonomyFieldType' || this.Type === 'User' || this.Type === 'Choice') {
            return this.Mult;
        }

        return null;
    }

    // eslint-disable-line
    private get fillIn(): boolean {
        
        if(this.Type === 'Choice') {
            return this.FillInChoice;
        }

        return null;
    }

    // eslint-disable-line
    private get format(): string {
        
        const allowedTypes = this.allAllowedFormatsForTypes[this.Type];

        if(isNullOrEmpty(allowedTypes)) {
            return null;
        }

        if(!allowedTypes.Contains(a => a === this.Format)) {
            return null;
        }

        return this.Format;
    }

    // eslint-disable-line
    private get userSelectionMode(): string {
        if(this.Type !== 'User') {
            return null;
        }

        return this.UserSelectionMode;
    }

    // eslint-disable-line
    private get richText(): boolean {
        if(this.Type == "Note") {
            return this.RichText;
        }

        return null;
    }

     // eslint-disable-line
     private get lookupList(): string {
        
        if(this.Type !== 'Lookup') {
            return null;
        }

        return this.List;
    }

    // eslint-disable-line
    private get showField(): string {
        
        if(this.Type !== 'Lookup') {
            return null;
        }

        return this.ShowField;
    }

    public get AdditionalField(): IField {
        if(this.Type !== 'TaxonomyFieldType') {
            return null;
        }

        const noteFieldForTaxonomy = new Field();
        noteFieldForTaxonomy.Type = 'Note';
        noteFieldForTaxonomy.DisplayName = `${this.Name}TaxHTField0`;
        noteFieldForTaxonomy.Name = `${this.Name}TaxHTField0`;
        noteFieldForTaxonomy.Required = false;
        noteFieldForTaxonomy.Hidden = true;
        //TODO ==> ShowInViewForms = false CanToggleHidden = true

        return noteFieldForTaxonomy;
    }

    protected attributeMapper: AttributeMapper[] = [
        { ownPropertyName: 'ID', elementName: 'ID' },
        { ownPropertyName: 'typeInXml', elementName: 'Type' },
        { ownPropertyName: 'Name', elementName: 'Name' },
        { ownPropertyName: 'Name', elementName: 'StaticName' },
        { ownPropertyName: 'DisplayName', elementName: 'DisplayName' },
        { ownPropertyName: 'Description', elementName: 'Description' },
        { ownPropertyName: 'Group', elementName: 'Group' },
        { ownPropertyName: 'Required', elementName: 'Required' },
        { ownPropertyName: 'Hidden', elementName: 'Hidden' },
        { ownPropertyName: 'Indexed', elementName: 'Indexed' },
        { ownPropertyName: 'multiAttribute', elementName: 'Mult' },
        { ownPropertyName: 'fillIn', elementName: 'FillInChoice' },
        { ownPropertyName: 'format', elementName: 'Format' },
        { ownPropertyName: 'userSelectionMode', elementName: 'UserSelectionMode' },
        { ownPropertyName: 'richText', elementName: 'RichText' },
        { ownPropertyName: 'RichTextMode', elementName: 'RichTextMode' },
        { ownPropertyName: 'lookupList', elementName: 'List' },
        { ownPropertyName: 'showField', elementName: 'ShowField' },
    ];

    constructor() {
        super();
        this.ID = '{' + Guid.newGuid().toString().toUpperCase() + '}';
    }

    public toElement(rootDocument: XMLDocument): Element {
        const element = super.toElement(rootDocument);
        
        if (this.Type == 'Choice') {
            const choices = rootDocument.createElement("CHOICES");
            
            if(!isNullOrEmpty(this.Choices)) {
                this.Choices.forEach((choiceValue: string) => {
                    const choice = rootDocument.createElement("CHOICE");
                    choice.innerHTML = choiceValue;
                    choices.appendChild(choice);
                });
            }

            element.appendChild(choices);
        }

        if(!isNullOrEmpty(this.DefaultValue)) {
            const defaultValue = rootDocument.createElement("Default");
            defaultValue.innerHTML = this.DefaultValue;
            element.appendChild(defaultValue);
        }
        
        return element;
    }


}