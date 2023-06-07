import * as React from 'react';
import styles from '../PnPTemplateGenerator.module.scss';
import { IBaseGeneratorComponentProps } from './IBaseGeneratorComponentProps';
import GenratorCommandBar from '../GeneratorCommandBar';
import { Dialog, DialogFooter, DefaultButton, PrimaryButton, IDropdownOption, Dropdown, TextField, DetailsList, IColumn, SelectionMode, Selection } from '@fluentui/react';
import { isFunction, isNullOrEmpty } from '@spfxappdev/utility';
import { IList, List } from '../../../../models';
import { cloneDeep } from '@microsoft/sp-lodash-subset';

export interface IListGeneratorProps extends IBaseGeneratorComponentProps {
}

interface IListGeneratorState {
    columns: IColumn[];
    items: IList[];
    commandbarItems: {
        isNewButtonDisabled: boolean;
        isEditButtonDisabled: boolean;
        isDeleteButtonDisabled: boolean;
    };
    showAddOrEditDialog: boolean;
    isAddOrUpdateButtonDisabled: boolean;
    currentList?: List;
}

export default class ListGenerator extends React.Component<IListGeneratorProps, IListGeneratorState> {

    private isAddNewMode: boolean = true;

    private selection: Selection;

    private get listTemplateTypeOptions(): IDropdownOption[] {
        let allAvailableListTypes: IDropdownOption[] = [
            { key: '100', text: 'List' },
            { key: '101', text: 'Document Library' },
            { key: '104', text: 'Announcements' },
            { key: '106', text: 'Calendar' },
            { key: '107', text: 'Tasks' }          
        ];

        return allAvailableListTypes;
    }

    public state: IListGeneratorState = {
        commandbarItems: {
            isNewButtonDisabled: false,
            isEditButtonDisabled: true,
            isDeleteButtonDisabled: true
        },
        showAddOrEditDialog: false,
        isAddOrUpdateButtonDisabled: true,
        columns: [{
            key: 'titleColumn',
            name: 'Title',
            fieldName: 'Title',
            minWidth: 100
        },
        {
            key: 'urlColumn',
            name: 'Url',
            fieldName: 'Url',
            minWidth: 100
        },
        {
            key: 'descriptionColumn',
            name: 'Description',
            fieldName: 'Description',
            minWidth: 100
        },
        {
            key: 'templateTypeColumn',
            name: 'Template Type',
            fieldName: 'TemplateType',
            minWidth: 100
        }
        ],
        items: cloneDeep(this.props.pnpTemplateGeneratorService.pnpTemplate.contentTypes),
    };

    constructor(props: IListGeneratorProps) {
        super(props);

        this.selection = new Selection({
            onSelectionChanged: () => {
                this.onSelectionChanged();
            },
        });
    }
    
    public render(): React.ReactElement<IListGeneratorProps> {
        return (<div className={styles.pnpTemplateGenerator}>
            <GenratorCommandBar
                isNewButtonDisabled={this.state.commandbarItems.isNewButtonDisabled}
                isEditButtonDisabled={this.state.commandbarItems.isEditButtonDisabled}
                isDeleteButtonDisabled={this.state.commandbarItems.isDeleteButtonDisabled}
                onNewButtonClick={() => { this.onAddNewListButtonClick(); }}
                onEditButtonClick={() => { this.onEditListButtonClick(); }}
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

        const { currentList } = this.state;

        return <Dialog
        hidden={!this.state.showAddOrEditDialog}
        >

            <Dropdown 
                options={this.listTemplateTypeOptions}
                selectedKey={currentList.TemplateType}
                label='Template Type'
                required={true}
                onChange={(ev, option: IDropdownOption) => {
                    const list = cloneDeep(this.state.currentList);
                    list.TemplateType = option.key.toString();

                    this.setState({
                        currentList: list,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(list)
                    });
                }}
            />

            <TextField 
                label='Title' 
                required={true} 
                defaultValue={currentList.Title}
                onChange={(ev: any, newValue: string) => {
                    const list = cloneDeep(this.state.currentList);
                    list.Title = newValue;

                    this.setState({
                        currentList: list,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(list)
                    });
                }}
            />

            <TextField 
                label='Url' 
                required={true} 
                defaultValue={currentList.Url}
                prefix={currentList.TemplateType == '101' ? '{weburl}/' : '{weburl}/Lists/'}
                onChange={(ev: any, newValue: string) => {
                    const list = cloneDeep(this.state.currentList);
                    list.Url = newValue;

                    this.setState({
                        currentList: list,
                        isAddOrUpdateButtonDisabled: this.isAddOrUpdateButtonDisabled(list)
                    });
                }}
            />

            <DialogFooter>
                <PrimaryButton 
                text={this.isAddNewMode?'Erstellen':'Speichern'} 
                disabled={this.state.isAddOrUpdateButtonDisabled}
                onClick={() => { 

                    if(this.isAddNewMode) {
                        this.onListAdded();
                        return;
                    }

                    this.onListUpdated();
                }} />
                <DefaultButton text='Abbrechen' onClick={() => { this.onAddOrEditDialogDismiss();  }} />
            </DialogFooter>
        </Dialog>;
    }

    private onAddOrEditDialogDismiss(): void {

        this.setState({
            showAddOrEditDialog: false,
            currentList: null,
            isAddOrUpdateButtonDisabled: false
        });

    }

    private onAddNewListButtonClick(): void {
        this.isAddNewMode = true;

        // const list = new List();
        
        // list.Title = 'Test';
        // list.Url = 'Lists/Test';
        // list.ContentTypeRefIds = [this.props.pnpTemplateGeneratorService.pnpTemplate.contentTypes[0].ID];

        // this.props.pnpTemplateGeneratorService.pnpTemplate.lists.push(list);

        this.setState({
            showAddOrEditDialog: true,
            currentList: new List()
        });
        
    }

    private onListAdded(): void {

        const list = cloneDeep(this.state.currentList);
        this.props.pnpTemplateGeneratorService.pnpTemplate.lists.push(list);
        const lists = cloneDeep(this.props.pnpTemplateGeneratorService.pnpTemplate.lists);

        this.setState({
            items: lists
        });

        this.onAddOrEditDialogDismiss();

        if(isFunction(this.props.onChange)) {
            this.props.onChange();
        }
    }

    private onEditListButtonClick(): void {
        this.isAddNewMode = false;

        this.setState({
            showAddOrEditDialog: true,
            currentList: cloneDeep(this.selection.getSelection()[0] as List),
            isAddOrUpdateButtonDisabled: false
        });
    }

    private onListUpdated(): void {

        const itemIndex = this.props.pnpTemplateGeneratorService.pnpTemplate.lists.IndexOf(i => i.UniqueId === this.state.currentList.UniqueId);

        this.props.pnpTemplateGeneratorService.pnpTemplate.lists[itemIndex] = cloneDeep(this.state.currentList);

        const lists = cloneDeep(this.props.pnpTemplateGeneratorService.pnpTemplate.lists);

        this.setState({
            items: lists
        });

        this.onAddOrEditDialogDismiss();

        if(isFunction(this.props.onChange)) {
            this.props.onChange();
        }

    }

    private isAddOrUpdateButtonDisabled(list: IList): boolean {
        
        if(isNullOrEmpty(list.Url)) {
            return true;
        }

        if(isNullOrEmpty(list.Title)) {
            return true;
        }

        return false;
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

}