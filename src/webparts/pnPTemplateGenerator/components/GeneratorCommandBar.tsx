import * as React from 'react';
import { ICommandBarItemProps, CommandBar } from '@fluentui/react';

export interface IGenratorCommandBarProps {
    isNewButtonDisabled: boolean;
    isEditButtonDisabled: boolean;
    isDeleteButtonDisabled: boolean;
    onNewButtonClick(): void;
    onEditButtonClick(): void;
    onDeleteButtonClick(): void;
}

interface IGenratorCommandBarState {
}

export default class GenratorCommandBar extends React.Component<IGenratorCommandBarProps, IGenratorCommandBarState> {

    public state: IGenratorCommandBarState = {
    };
    
    public render(): React.ReactElement<IGenratorCommandBarProps> {

        const commandBarItems: ICommandBarItemProps[] = [{
            key: 'newField',
            text: 'New',
            iconProps: { iconName: 'Add' },
            disabled: this.props.isNewButtonDisabled,
            onClick: () => {
                this.props.onNewButtonClick();
            }
        },
        {
            key: 'editField',
            text: 'Edit',
            iconProps: { iconName: 'Edit' },
            disabled: this.props.isEditButtonDisabled,
            onClick: () => {
                this.props.onEditButtonClick();
            }
        },
        {
            key: 'deleteField',
            text: 'Delete',
            iconProps: { iconName: 'Delete' },
            disabled: this.props.isDeleteButtonDisabled,
            onClick: () => {
                this.props.onDeleteButtonClick();
            }
        }
        ];
        return (<CommandBar items={commandBarItems} />)
    }
}