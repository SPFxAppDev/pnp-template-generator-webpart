export interface IBaseTemplateModel {
    elementName: string;
    toElement(rootDocument: XMLDocument): Element;
}