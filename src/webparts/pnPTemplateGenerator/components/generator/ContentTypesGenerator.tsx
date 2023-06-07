import * as React from 'react';
import styles from '../PnPTemplateGenerator.module.scss';
import { IColumn, DetailsList, SelectionMode, Selection, Dialog, DialogFooter, DefaultButton, PrimaryButton, TextField, Dropdown, IDropdownOption, SelectableOptionMenuItemType, Toggle } from '@fluentui/react';
import { ContentType, IContentType, IField } from '../../../../models';
import { IBaseGeneratorComponentProps } from './IBaseGeneratorComponentProps';
import { isFunction, isNullOrEmpty } from '@spfxappdev/utility';
import { cloneDeep } from '@microsoft/sp-lodash-subset';
import GenratorCommandBar from '../GeneratorCommandBar';

export interface IContentTypesGeneratorProps extends IBaseGeneratorComponentProps {
    
}

interface IContentTypesGeneratorState {
    columns: IColumn[];
    items: IContentType[];
    commandbarItems: {
        isNewButtonDisabled: boolean;
        isEditButtonDisabled: boolean;
        isDeleteButtonDisabled: boolean;
    };
    showAddOrEditDialog: boolean;
    currentCT?: ContentType;
    isAddOrUpdateButtonDisabled: boolean;
}

export default class ContentTypesGenerator extends React.Component<IContentTypesGeneratorProps, IContentTypesGeneratorState> {

    private selection: Selection;

    private isAddNewMode: boolean = true;

    private get parentContentTypeOptions(): IDropdownOption[] {
        let allAvailableParentContentTypes: IDropdownOption[] = [
            { key: 'HeaderSystemContentTypes', text: "Default Content Types", itemType: SelectableOptionMenuItemType.Header },
            { key: '0x01', text: 'Item' },
            { key: '0x0101', text: 'Document' },
            { key: '0x0102', text: 'Event' },
            { key: '0x0104', text: 'Announcement' },
            { key: '0x0108', text: 'Task' },
            { key: '0x0120', text: 'Folder' }            
        ];

        let { contentTypes } = this.props.pnpTemplateGeneratorService.pnpTemplate;

        if(!isNullOrEmpty(contentTypes)) {
            contentTypes = contentTypes.Where(ct => !ct.ID.StartsWith(this.state.currentCT.ID));
        }

        if(!isNullOrEmpty(contentTypes)) {

            allAvailableParentContentTypes.push({ key: 'HeaderCustomContentTypes', text: "Custom Content Types", itemType: SelectableOptionMenuItemType.Header });

            const customContentTypeOptions: IDropdownOption[] = contentTypes.map((contentType: ContentType): IDropdownOption => {
                return { key: contentType.ID, text: contentType.Name }
            });

            allAvailableParentContentTypes = allAvailableParentContentTypes.concat(customContentTypeOptions);
        }

        return allAvailableParentContentTypes;
    }

    private get fieldRefsTypeOptions(): IDropdownOption[] {
        let allAvailableFields: IDropdownOption[] = [];

        const { siteColumns } = this.props.pnpTemplateGeneratorService.pnpTemplate;

        if(isNullOrEmpty(siteColumns)) {
            return allAvailableFields;
        }

        siteColumns.forEach((field: IField) => {
            allAvailableFields.push({ key: field.ID, text: `${field.DisplayName} (${field.Name})` })
        });

        return allAvailableFields;
    }

    public state: IContentTypesGeneratorState = {
        showAddOrEditDialog: false,
        isAddOrUpdateButtonDisabled: true,
        columns: [{
            key: 'displayNameColumn',
            name: 'Name',
            fieldName: 'Name',
            minWidth: 100
        },
        {
            key: 'descriptionColumn',
            name: 'Description',
            fieldName: 'Description',
            minWidth: 100
        },
        {
            key: 'groupColumn',
            name: 'Group',
            fieldName: 'Group',
            minWidth: 100
        }
        ],
        items: cloneDeep(this.props.pnpTemplateGeneratorService.pnpTemplate.contentTypes),
        commandbarItems: {
            isNewButtonDisabled: false,
            isEditButtonDisabled: true,
            isDeleteButtonDisabled: true
        }
    };

    constructor(props: IContentTypesGeneratorProps) {
        super(props);

        this.selection = new Selection({
            onSelectionChanged: () => {
                this.onSelectionChanged();
            },
        });
    }
    
    public render(): React.ReactElement<IContentTypesGeneratorProps> {
        return (<div className={styles.pnpTemplateGenerator}>
            
            <GenratorCommandBar
                isNewButtonDisabled={this.state.commandbarItems.isNewButtonDisabled}
                isEditButtonDisabled={this.state.commandbarItems.isEditButtonDisabled}
                isDeleteButtonDisabled={this.state.commandbarItems.isDeleteButtonDisabled}
                onNewButtonClick={() => { this.onAddNewContentTypeButtonClick(); }}
                onEditButtonClick={() => { this.onEditContentTypeButtonClick(); }}
                onDeleteButtonClick={() => { console.log("DeleteButton Clicked"); }}
            />

            <DetailsList 
                items={this.state.items}
                columns={this.state.columns}
                selection={this.selection}
                selectionMode={SelectionMode.multiple}
            />

            {this.state.showAddOrEditDialog && this.renderAddOrEditDialog()}
        </div>);
    }

    private renderAddOrEditDialog(): JSX.Element {

        const { currentCT } = this.state;

        return <Dialog
        hidden={!this.state.showAddOrEditDialog}
        >

            <TextField label='ID' disabled={true} value={currentCT.ID} />
            
            <Dropdown 
                options={this.parentContentTypeOptions}
                selectedKey={currentCT.ParentID}
                label='Parent Content Type'
                required={true}
                onChange={(ev, option: IDropdownOption) => {
                    const contentType = cloneDeep(this.state.currentCT);
                    contentType.ParentID = option.key.toString();

                    this.setState({
                        currentCT: contentType,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(contentType)
                    });
                }}
            />

            <TextField 
                label='Name' 
                required={true} 
                defaultValue={currentCT.Name}
                onChange={(ev: any, newValue: string) => {
                    const contentType = cloneDeep(this.state.currentCT);
                    contentType.Name = newValue;

                    this.setState({
                        currentCT: contentType,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(contentType)
                    });
                }}
            />

            <Dropdown 
                options={this.fieldRefsTypeOptions}
                selectedKeys={currentCT.FieldRefIds}
                label='Fields'
                disabled={isNullOrEmpty(currentCT.FieldRefIds)}
                multiSelect={true}
                onChange={(ev, option: IDropdownOption) => {
                    const contentType = cloneDeep(this.state.currentCT);
                    
                    contentType.FieldRefIds = option.selected ? [...contentType.FieldRefIds, option.key as string] : contentType.FieldRefIds.filter(key => key !== option.key);

                    this.setState({
                        currentCT: contentType,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(contentType)
                    });
                }}
            />

            <TextField 
                label='Description'
                multiline={true}
                maxLength={250}
                defaultValue={currentCT.Description}
                onChange={(ev: any, newValue: string) => {
                    const contentType = cloneDeep(this.state.currentCT);
                    contentType.Description = newValue;

                    this.setState({
                        currentCT: contentType,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(contentType)
                    });
                }}
            />

            <TextField 
                label='Group' 
                defaultValue={currentCT.Group}
                onChange={(ev: any, newValue: string) => {
                    const contentType = cloneDeep(this.state.currentCT);
                    contentType.Group = newValue;

                    this.setState({
                        currentCT: contentType,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(contentType)
                    });
                }}
            />

            <Toggle
                label='Hidden'
                defaultChecked={currentCT.Hidden}
                offText='No'
                onText='Yes'
                onChange={(ev: any, checked) => {
                    const contentType = cloneDeep(this.state.currentCT);
                    contentType.Hidden = checked;

                    this.setState({
                        currentCT: contentType,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(contentType)
                    });
                }}
            />

            

            <DialogFooter>
                <PrimaryButton 
                text={this.isAddNewMode?'Erstellen':'Speichern'} 
                disabled={this.state.isAddOrUpdateButtonDisabled}
                onClick={() => { 

                    if(this.isAddNewMode) {
                        this.onContentTypeAdded();
                        return;
                    }

                    this.onContentTypeUpdated();
                }} />
                <DefaultButton text='Abbrechen' onClick={() => { this.onAddOrEditDialogDismiss();  }} />
            </DialogFooter>
        </Dialog>;
    }

    private onAddOrEditDialogDismiss(): void {

        this.setState({
            showAddOrEditDialog: false,
            currentCT: null,
            isAddOrUpdateButtonDisabled: true
        });

    }

    private onAddNewContentTypeButtonClick(): void {
        this.isAddNewMode = true;

        this.setState({
            showAddOrEditDialog: true,
            currentCT: new ContentType()
        });
    }

    private onContentTypeAdded(): void {
        const contentType = cloneDeep(this.state.currentCT);
        this.props.pnpTemplateGeneratorService.pnpTemplate.contentTypes.push(contentType);
        const contentTypes = cloneDeep(this.props.pnpTemplateGeneratorService.pnpTemplate.contentTypes);

        this.setState({
            items: contentTypes
        });

        this.onAddOrEditDialogDismiss();

        if(isFunction(this.props.onChange)) {
            this.props.onChange();
        }
    }

    private onEditContentTypeButtonClick(): void {
        this.isAddNewMode = false;

        this.setState({
            showAddOrEditDialog: true,
            currentCT: cloneDeep(this.selection.getSelection()[0] as ContentType),
            isAddOrUpdateButtonDisabled: false
        });
    }

    private onContentTypeUpdated(): void {

        const itemIndex = this.props.pnpTemplateGeneratorService.pnpTemplate.contentTypes.IndexOf(i => i.getInternalIdentifier() === this.state.currentCT.getInternalIdentifier());

        this.props.pnpTemplateGeneratorService.pnpTemplate.contentTypes[itemIndex] = cloneDeep(this.state.currentCT);

        const contentTypes = cloneDeep(this.props.pnpTemplateGeneratorService.pnpTemplate.contentTypes);

        this.setState({
            items: contentTypes
        });

        this.onAddOrEditDialogDismiss();

        if(isFunction(this.props.onChange)) {
            this.props.onChange();
        }

    }

    private onSelectionChanged(): void {
        const count = this.selection.getSelectedCount();
        const commandbarItems = {...this.state.commandbarItems};

        if(count == 0) {
            commandbarItems.isNewButtonDisabled = false;
            commandbarItems.isEditButtonDisabled = true;
            commandbarItems.isDeleteButtonDisabled = true;
        }

        if(count >= 1) {
            commandbarItems.isNewButtonDisabled = true;
            commandbarItems.isEditButtonDisabled = true;
            commandbarItems.isDeleteButtonDisabled = false;
        }

        if(count == 1) {
            commandbarItems.isEditButtonDisabled = false;
        }
        
        this.setState({
            commandbarItems: commandbarItems
        });
    }

    private isAddOrUpdateButtonDisabled(contentType: ContentType): boolean {

        if(isNullOrEmpty(contentType.ParentID)) {
            return true;
        }

        if(isNullOrEmpty(contentType.Name)) {
            return true;
        }

        return false;
    }
}