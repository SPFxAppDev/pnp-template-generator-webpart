import * as React from 'react';
import styles from '../PnPTemplateGenerator.module.scss';
//, DetailsListLayoutMode, Selection, 
import { IColumn, DetailsList, SelectionMode, Selection, Dialog, DialogFooter, DefaultButton, PrimaryButton, TextField, Dropdown, IDropdownOption, Toggle, ChoiceGroup } from '@fluentui/react';
import { Field, IField } from '../../../../models';
import { IBaseGeneratorComponentProps } from './IBaseGeneratorComponentProps';
import { isFunction, isNullOrEmpty } from '@spfxappdev/utility';
import { cloneDeep } from '@microsoft/sp-lodash-subset';
import GenratorCommandBar from '../GeneratorCommandBar';

export interface ISiteColumnsGeneratorProps extends IBaseGeneratorComponentProps {
}

interface ISiteColumnsGeneratorState {
    columns: IColumn[];
    items: IField[];
    commandbarItems: {
        isNewButtonDisabled: boolean;
        isEditButtonDisabled: boolean;
        isDeleteButtonDisabled: boolean;
    };
    showAddOrEditDialog: boolean;
    currentField?: Field;
    isAddOrUpdateButtonDisabled: boolean;
}

const fieldTypeOptions: IDropdownOption[] = [
    { key: 'Text', text: 'Single line of Text' },
    { key: 'Note', text: 'Multiple lines of text' },
    { key: 'Choice', text: 'Choice' },
    { key: 'Number', text: 'Number' },
    { key: 'DateTime', text: 'Date and Time' },
    { key: 'Lookup', text: 'Lookup' },
    { key: 'Boolean', text: 'Yes / No' },
    { key: 'User', text: 'Person or Group' },
    { key: 'URL', text: 'Hyperlink or Picture' },
    { key: 'TaxonomyFieldType', text: 'Managed Metadata' },
  ];

export default class SiteColumnsGenerator extends React.Component<ISiteColumnsGeneratorProps, ISiteColumnsGeneratorState> {

    private selection: Selection;

    private isAddNewMode: boolean = true;

    public state: ISiteColumnsGeneratorState = {
        showAddOrEditDialog: false,
        isAddOrUpdateButtonDisabled: true,
        columns: [{
            key: 'displayNameColumn',
            name: 'Display Name',
            fieldName: 'DisplayName',
            minWidth: 100
        },
        {
            key: 'internalNameColumn',
            name: 'Internal Name',
            fieldName: 'Name',
            minWidth: 100
        },
        {
            key: 'typeColumn',
            name: 'Type',
            fieldName: 'Type',
            minWidth: 100
        }
        ],
        items: cloneDeep(this.props.pnpTemplateGeneratorService.pnpTemplate.siteColumns),
        commandbarItems: {
            isNewButtonDisabled: false,
            isEditButtonDisabled: true,
            isDeleteButtonDisabled: true
        }
    };

    constructor(props: ISiteColumnsGeneratorProps) {
        super(props);

        this.selection = new Selection({
            onSelectionChanged: () => {
                this.onSelectionChanged();
            },
        });
    }
    
    public render(): React.ReactElement<ISiteColumnsGeneratorProps> {

        return (<div className={styles.pnpTemplateGenerator}>

            <GenratorCommandBar
                isNewButtonDisabled={this.state.commandbarItems.isNewButtonDisabled}
                isEditButtonDisabled={this.state.commandbarItems.isEditButtonDisabled}
                isDeleteButtonDisabled={this.state.commandbarItems.isDeleteButtonDisabled}
                onNewButtonClick={() => { this.onAddNewFieldButtonClick(); }}
                onEditButtonClick={() => { this.onEditFieldButtonClick(); }}
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

        const { currentField } = this.state;

        return <Dialog
        hidden={!this.state.showAddOrEditDialog}
        >
            <TextField label='ID' disabled={true} defaultValue={currentField.ID} />

            <TextField 
                label='Display Name' 
                required={true} 
                defaultValue={currentField.DisplayName}
                onChange={(ev: any, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.DisplayName = newValue;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <TextField 
                label='Internal Name' 
                required={true} 
                defaultValue={currentField.Name}
                onChange={(ev: any, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Name = newValue;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <Dropdown
                label='Type'
                required={true}
                options={fieldTypeOptions} 
                defaultSelectedKey={currentField.Type}
                onChange={(ev: any, option: IDropdownOption) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Type = option.key as string;

                    if(field.Type === 'Boolean') {
                        field.Required = false;
                        field.DefaultValue = '1';
                    }

                    if(field.Type === 'URL') {
                        field.Format = 'Hyperlink';
                    }

                    if(field.Type === 'User') {
                        field.UserSelectionMode = 'PeopleOnly';
                    }

                    if(field.Type === 'Note') {
                        field.RichText = true;
                    }

                    if(field.Type === 'DateTime') {
                        field.Format = 'DateOnly';
                    }

                    if(field.Type === 'Lookup') {
                        field.Format = 'Dropdown';
                        field.ShowField = 'Title';
                    }

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            
            <Toggle
                label='Require that this column contains information'
                defaultChecked={currentField.Required}
                offText='No'
                onText='Yes'
                disabled={currentField.Type === 'Boolean'}
                onChange={(ev: any, checked) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Required = checked;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <TextField 
                label='Description'
                multiline={true}
                maxLength={250}
                defaultValue={currentField.Description}
                onChange={(ev: any, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Description = newValue;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <TextField 
                label='Group'
                defaultValue={currentField.Group}
                onChange={(ev: any, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Group = newValue;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <Toggle
                label='Hidden'
                defaultChecked={currentField.Hidden}
                offText='No'
                onText='Yes'
                onChange={(ev: any, checked) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Hidden = checked;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            {this.renderTypeSpecificSettings(currentField)}

            <DialogFooter>
                <PrimaryButton 
                text={this.isAddNewMode?'Erstellen':'Speichern'} 
                disabled={this.state.isAddOrUpdateButtonDisabled}
                onClick={() => { 

                    if(this.isAddNewMode) {
                        this.onFieldAdded();
                        return;
                    }

                    this.onFieldUpdated();
                }} />
                <DefaultButton text='Abbrechen' onClick={() => { this.onAddOrEditDialogDismiss();  }} />
            </DialogFooter>
        </Dialog>;
    }

    private renderTypeSpecificSettings(currentField: Field): JSX.Element {

        const defaultTextValueAllowedTypes = ['Choice', 'Number', 'Text'];
        let renderDefaultValueTextBox: boolean = false;
        if(defaultTextValueAllowedTypes.Contains(t => t === currentField.Type)) {
            renderDefaultValueTextBox = true;
        }


        return <>
            {currentField.Type === 'Choice' && this.renderDropDownSettings(currentField)}
            {currentField.Type === 'TaxonomyFieldType' && this.renderTaxonomySettings(currentField)}
            {currentField.Type === 'Note' && this.renderNoteSettings(currentField)}
            {currentField.Type === 'DateTime' && this.renderDateTimeSettings(currentField)}
            {currentField.Type === 'Lookup' && this.renderLookupSettings(currentField)}
            {currentField.Type === 'Boolean' && this.renderBooleanSettings(currentField)}
            {currentField.Type === 'User' && this.renderUserSettings(currentField)}
            {currentField.Type === 'URL' && this.renderUrlSettings(currentField)}

            {renderDefaultValueTextBox && 
            <TextField
                label='Default Value'
                defaultValue={currentField.DefaultValue}
                onChange={(ev, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.DefaultValue = newValue;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />}
        </>;
    }

    private renderDropDownSettings(currentField: Field): JSX.Element {
        return (<>
            <TextField
                label='Type each choice on a separate line'
                required={true}
                multiline={true}
                defaultValue={isNullOrEmpty(currentField.Choices) ? '' : currentField.Choices.join('\n')}
                onChange={(ev, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Choices = isNullOrEmpty(newValue) ? null : newValue.split('\n');

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <Toggle
                label='Allow "Fill-in" choices'
                defaultChecked={currentField.FillInChoice}
                offText='No'
                onText='Yes'
                onChange={(ev: any, checked) => {
                    const field = cloneDeep(this.state.currentField);
                    field.FillInChoice = checked;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <ChoiceGroup 
                defaultSelectedKey={currentField.Mult ? 'Checkboxes' : isNullOrEmpty(currentField.Format) ? 'Dropdown' : currentField.Format} 
                options={[
                    { key: 'Dropdown', text: 'Drop-Down Menu' },
                    { key: 'RadioButtons', text: 'Radio Buttons' },
                    { key: 'Checkboxes', text: 'Checkboxes (allow multiple selections)' }
                ]} 
                onChange={(ev, option) => {
                    const field = cloneDeep(this.state.currentField);

                    if(option.key == 'Checkboxes') {
                        field.Mult = true;
                        field.Format = undefined;
                    }
                    else {
                        field.Mult = undefined;
                        field.Format = option.key;
                    }

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });

                }} 
                label="Display type"
            />
        </>);
    }

    private renderTaxonomySettings(currentField: Field): JSX.Element {
        return <>
            <Toggle
                label='Allow mulitple values'
                defaultChecked={currentField.Mult}
                offText='No'
                onText='Yes'
                onChange={(ev: any, checked) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Mult = checked;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <TextField 
                label='Termgroup Name'
                defaultValue={currentField.TermGroupName}
                onChange={(ev: any, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.TermGroupName = newValue;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <TextField 
                label='Termgset Name'
                defaultValue={currentField.TermSetName}
                onChange={(ev: any, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.TermSetName = newValue;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />
        </>;
    }

    private renderNoteSettings(currentField: Field): JSX.Element {
        const formatOptions = [
            { key: 'PlainText', text: 'Plain text' },
            { key: 'RichText', text: 'Enhanced rich text (Rich text with pictures, tables, and hyperlinks)' }
        ];

        return (<>
            <ChoiceGroup 
                defaultSelectedKey={currentField.RichText ? 'RichText' : 'PlainText'}
                options={formatOptions} 
                onChange={(ev, option) => {
                    const field = cloneDeep(this.state.currentField);
                    field.RichText = option.key == 'RichText';


                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });

                }} 
                label="Display type"
            />
        </>);
    }

    private renderDateTimeSettings(currentField: Field): JSX.Element {
        const formatOptions = [
            { key: 'DateOnly', text: 'Date only' },
            { key: 'DateTime', text: 'Date and time' }
        ];

        return (<>
            <ChoiceGroup 
                defaultSelectedKey={currentField.Format}
                options={formatOptions} 
                onChange={(ev, option) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Format = option.key.toString();

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });

                }}
                label="Format"
            />
        </>);
    }

    private renderLookupSettings(currentField: Field): JSX.Element {
        return (<>
            <TextField
                label='Get information from this List Url'
                required={true}
                description='Add the site relative list url i. e. Lists/MyListUrl or for Libraries i. e. MyBibUrl'
                defaultValue={currentField.List}
                onChange={(ev, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.List = newValue;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />

            <TextField
                label='Show information from this internal field name'
                required={true}
                defaultValue={currentField.ShowField}
                onChange={(ev, newValue: string) => {
                    const field = cloneDeep(this.state.currentField);
                    field.ShowField = newValue;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />
        </>);
    }

    private renderBooleanSettings(currentField: Field): JSX.Element {
        const defaultValueOptions = [
            { key: '1', text: 'Yes' },
            { key: '0', text: 'No' }
        ];

        return (<>
            <ChoiceGroup 
                defaultSelectedKey={currentField.DefaultValue}
                options={defaultValueOptions} 
                onChange={(ev, option) => {
                    const field = cloneDeep(this.state.currentField);
                    field.DefaultValue = option.key.toString();

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });

                }}
                label="Default value"
            />
        </>);
    }

    private renderUserSettings(currentField: Field): JSX.Element {
        const userSelectionModeOptions = [
            { key: 'PeopleOnly', text: 'People only' },
            { key: 'PeopleAndGroups', text: 'People and Groups' }
        ];

        return (<>
            <ChoiceGroup 
                defaultSelectedKey={currentField.UserSelectionMode}
                options={userSelectionModeOptions} 
                onChange={(ev, option) => {
                    const field = cloneDeep(this.state.currentField);
                    field.UserSelectionMode = option.key;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });

                }} 
                label="Allow selection of"
            />

            <Toggle
                label='Allow mulitple selections'
                defaultChecked={currentField.Mult}
                offText='No'
                onText='Yes'
                onChange={(ev: any, checked) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Mult = checked;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });
                }}
            />
        </>);
    }

    private renderUrlSettings(currentField: Field): JSX.Element {

        const formatOptions = [
            { key: 'Hyperlink', text: 'Hyperlink' },
            { key: 'Image', text: 'Picture' }
        ];

        return (<>
            <ChoiceGroup 
                defaultSelectedKey={currentField.Format}
                options={formatOptions} 
                onChange={(ev, option) => {
                    const field = cloneDeep(this.state.currentField);
                    field.Format = option.key;

                    this.setState({
                        currentField: field,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(field)
                    });

                }} 
                label="Display type"
            />
        </>);
    }

    private onAddOrEditDialogDismiss(): void {

        this.setState({
            showAddOrEditDialog: false,
            currentField: null,
            isAddOrUpdateButtonDisabled: true
        });

    }

    private onAddNewFieldButtonClick(): void {
        this.isAddNewMode = true;

        this.setState({
            showAddOrEditDialog: true,
            currentField: new Field()
        });
        
    }

    private onFieldAdded(): void {
        const field = cloneDeep(this.state.currentField);
        this.props.pnpTemplateGeneratorService.pnpTemplate.siteColumns.push(field);
        const fields = cloneDeep(this.props.pnpTemplateGeneratorService.pnpTemplate.siteColumns);

        this.setState({
            items: fields
        });

        this.onAddOrEditDialogDismiss();

        if(isFunction(this.props.onChange)) {
            this.props.onChange();
        }
    }

    private onEditFieldButtonClick(): void {
        this.isAddNewMode = false;

        this.setState({
            showAddOrEditDialog: true,
            currentField: cloneDeep(this.selection.getSelection()[0] as Field),
            isAddOrUpdateButtonDisabled: false
        });
    }

    private onFieldUpdated(): void {

        const fieldIndex = this.props.pnpTemplateGeneratorService.pnpTemplate.siteColumns.IndexOf(i => i.ID === this.state.currentField.ID);

        this.props.pnpTemplateGeneratorService.pnpTemplate.siteColumns[fieldIndex] = cloneDeep(this.state.currentField);

        const fields = cloneDeep(this.props.pnpTemplateGeneratorService.pnpTemplate.siteColumns);

        this.setState({
            items: fields
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

    private isAddOrUpdateButtonDisabled(field: Field): boolean {
        
        if(!isNullOrEmpty(field.Name) && this.isInternalNameAlreadyExist(field.Name)) {
            return true;
        }

        if(isNullOrEmpty(field.DisplayName)) {
            return true;
        }

        if(isNullOrEmpty(field.Name)) {
            return true;
        }

        if(isNullOrEmpty(field.Type)) {
            return true;
        }

        if(field.Type === 'Choice' && isNullOrEmpty(field.Choices)) {
            return true;
        }

        if(field.Type === 'Lookup' && (isNullOrEmpty(field.List) || isNullOrEmpty(field.ShowField))) {
            return true;
        }
    }

    private isInternalNameAlreadyExist(internalName: string): boolean {

        const { siteColumns } = this.props.pnpTemplateGeneratorService.pnpTemplate; 

        if(!this.isAddNewMode && (this.selection.getSelection()[0] as Field).Name === internalName) {
            return false;
        }

        return siteColumns.Contains(i => i.Name.Equals(internalName, true));

    }
}