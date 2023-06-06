import { IPnPTemplateGeneratorServiceService } from "../../../../services/PnPTemplateGeneratorService";

export interface IBaseGeneratorComponentProps {
    pnpTemplateGeneratorService: IPnPTemplateGeneratorServiceService;
    onChange?(): void;
}