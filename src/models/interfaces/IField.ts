import { IBaseTemplateModel } from "./IBaseTemplateModel";

export interface IField extends IBaseTemplateModel {
    ID: string;
    Type: string;
    Name: string;
    DisplayName: string;
    Description?: string;
    Group?: string;
    Required?: boolean;
    Hidden?: boolean;
    Format?: string;
    DefaultValue?: string;
    Indexed?: boolean;

    //Choice
    Choices?: string[];
    FillInChoice?: boolean;

    //Taxonomy + UserMulti
    Mult?: boolean;

    //Note
    RichText?: boolean;
    RichTextMode?: string;
    NumLines?: number;

    //Lookup
    List?: string;
    ShowField?: string;

    //Taxonomy
    SourceID?: string;
    TermGroupName?: string;
    TermSetName?: string;

    //User
    UserSelectionMode?: string;



}