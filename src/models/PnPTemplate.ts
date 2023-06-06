import { Field } from "./Field";


export class PnPTemplate {

    //TODO: Ggf. mit FieldCollection, ContentTypeCollection und ListCollection arbeiten ==> Die handlen dann intern add usw.
    public siteColumns: Field[] = [];

    public contentTypes: any[] = [];

    public lists: any[] = [];
    
}