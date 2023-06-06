import { isNullOrEmpty } from "@spfxappdev/utility";
import { IBaseTemplateModel } from "./interfaces";

export type AttributeMapper = { ownPropertyName: string; elementName: string; };

export abstract class ElementBase implements IBaseTemplateModel {
    public elementName: string;

    protected attributeMapper: AttributeMapper[] = [];

    public toElement(rootDocument: XMLDocument): Element {
        const element = rootDocument.createElement(this.elementName);

        if(!isNullOrEmpty(this.attributeMapper)) {

            this.attributeMapper.forEach((mapper: AttributeMapper) => {

                const self = this as any;
                const ownProp = self[mapper.ownPropertyName];

                if(!isNullOrEmpty(ownProp)) {

                    if(typeof ownProp === 'string') {
                        element.setAttribute(mapper.elementName, ownProp);
                    }
                    else if(typeof ownProp === 'boolean') {
                        element.setAttribute(mapper.elementName, ownProp ? 'TRUE' : 'FALSE');
                    }
                    else if(typeof ownProp === 'number') {
                        element.setAttribute(mapper.elementName, ownProp.toString());
                    }
                }

            });

        }

        return element;
    }
    
}