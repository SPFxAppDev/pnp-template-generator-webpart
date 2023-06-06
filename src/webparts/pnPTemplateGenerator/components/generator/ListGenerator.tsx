import * as React from 'react';
import styles from '../PnPTemplateGenerator.module.scss';
import { IBaseGeneratorComponentProps } from './IBaseGeneratorComponentProps';
import GenratorCommandBar from '../GeneratorCommandBar';
import { Dialog, DialogFooter, DefaultButton, PrimaryButton } from '@fluentui/react';
import { isFunction } from '@spfxappdev/utility';
import { IList, List } from '../../../../models';

export interface IListGeneratorProps extends IBaseGeneratorComponentProps {
}

interface IListGeneratorState {
    commandbarItems: {
        isNewButtonDisabled: boolean;
        isEditButtonDisabled: boolean;
        isDeleteButtonDisabled: boolean;
    };
    showAddOrEditDialog: boolean;
    isAddOrUpdateButtonDisabled: boolean;
}

export default class ListGenerator extends React.Component<IListGeneratorProps, IListGeneratorState> {

    private isAddNewMode: boolean = true;

    public state: IListGeneratorState = {
        commandbarItems: {
            isNewButtonDisabled: false,
            isEditButtonDisabled: true,
            isDeleteButtonDisabled: true
        },
        showAddOrEditDialog: false,
        isAddOrUpdateButtonDisabled: true,
    };
    
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
            
            {/* <CommandBar items={this.state.commandbarItems} />

            <DetailsList 
                items={this.state.items}
                columns={this.state.columns}
                selection={this.selection}
                selectionMode={SelectionMode.multiple}
            />*/}

            {this.state.showAddOrEditDialog && this.renderAddOrEditDialog()} 
        </div>);
    }

    private renderAddOrEditDialog(): JSX.Element {

        return <Dialog
        hidden={!this.state.showAddOrEditDialog}
        >

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
            // currentField: null,
            isAddOrUpdateButtonDisabled: false
        });

    }

    private onAddNewListButtonClick(): void {
        this.isAddNewMode = true;

        const list = new List();
        
        list.Title = 'Test';
        list.Url = 'Lists/Test';
        list.ContentTypeRefIds = [this.props.pnpTemplateGeneratorService.pnpTemplate.contentTypes[0].ID];

        this.props.pnpTemplateGeneratorService.pnpTemplate.lists.push(list);

        this.setState({
            showAddOrEditDialog: true,
            // currentField: new Field()
        });
        
    }

    private onListAdded(): void {

        this.onAddOrEditDialogDismiss();

        if(isFunction(this.props.onChange)) {
            this.props.onChange();
        }
    }

    private onEditListButtonClick(): void {
        this.isAddNewMode = false;

        this.setState({
            showAddOrEditDialog: true,
            // currentField: cloneDeep(this.selection.getSelection()[0] as Field),
            isAddOrUpdateButtonDisabled: false
        });
    }

    private onListUpdated(): void {

        this.onAddOrEditDialogDismiss();

        if(isFunction(this.props.onChange)) {
            this.props.onChange();
        }

    }

    private isAddOrUpdateButtonDisabled(list: IList): boolean {
        
        return false;
    }

}