import * as React from 'react';
import * as OfficeUI from 'office-ui-fabric-react';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import {
    DetailsList,
    DetailsListLayoutMode,
    Selection,
    SelectionMode,
    IColumn,
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';


const classNames = mergeStyleSets({
    fileIconHeaderIcon: {
        padding: 0,
        fontSize: '16px',
    },
    fileIconCell: {
        textAlign: 'center',
        selectors: {
            '&:before': {
                content: '.',
                display: 'inline-block',
                verticalAlign: 'middle',
                height: '100%',
                width: '0px',
                visibility: 'hidden',
            },
        },
    },
    fileIconImg: {
        verticalAlign: 'middle',
        maxHeight: '16px',
        maxWidth: '16px',
    },
    controlWrapper: {
        display: 'flex',
        flexWrap: 'wrap',
    },
    exampleToggle: {
        display: 'inline-block',
        marginBottom: '10px',
        marginRight: '30px',
    },
    selectionDetails: {
        marginBottom: '20px',
    },
});

export interface IDetailsListDocumentsExampleState {
    columns: IColumn[];
    selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
    announcedMessage?: string;
}

export interface IDocument {
    key: string;
    Title: string;
    ClientNumber: string;
}

export interface IProps {
    UserState: any;
    UserControlAction: any;
}

export class CLUserList extends React.Component<IProps> {
    private _selection = new Selection({
        onSelectionChanged: () => {
            this.setState({
                selectionDetails: this._getSelectionDetails(),
            });
        },
    });

    private _allItems: any;
    private columns: IColumn[] = [
        {
            key: 'column2',
            name: 'CL-TAX-WF',
            fieldName: 'Email',
            minWidth: 190,
            maxWidth: 190,
            isRowHeader: true,
            isResizable: true,
            data: 'string',
            isPadded: true,
        },
    ];

    public state = {
        columns: this.columns,
        selectionDetails: this._getSelectionDetails(),
        isModalSelection: true,
        isCompactMode: false,
        announcedMessage: undefined,
    };


    public componentDidMount() {
        this._allItems = this.props.UserState;
    }


    public render() {
        const { columns, isCompactMode, isModalSelection, announcedMessage } = this.state;
        const items = this.props.UserState;
        return (
            <Fabric>
                <div className={classNames.controlWrapper}>

                </div>
                {announcedMessage ? <Announced message={announcedMessage} /> : undefined}
                {isModalSelection ? (
                    <MarqueeSelection selection={this._selection}>
                        <DetailsList
                            items={items}
                            compact={isCompactMode}
                            columns={columns}
                            selectionMode={SelectionMode.multiple}
                            getKey={this._getKey}
                            setKey="multiple"
                            layoutMode={DetailsListLayoutMode.justified}
                            isHeaderVisible={true}
                            selection={this._selection}
                            selectionPreservedOnEmptyClick={true}
                            enterModalSelectionOnTouch={true}
                            ariaLabelForSelectionColumn="Toggle selection"
                            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                            checkButtonAriaLabel="Row checkbox"
                            onRenderItemColumn={this._onRenderItemColumn}

                        />
                    </MarqueeSelection>
                ) : (
                    <DetailsList
                        items={items}
                        compact={isCompactMode}
                        columns={columns}
                        selectionMode={SelectionMode.none}
                        getKey={this._getKey}
                        setKey="none"
                        layoutMode={DetailsListLayoutMode.justified}
                        isHeaderVisible={true}
                        onRenderItemColumn={this._onRenderItemColumn}
                    />
                )}
            </Fabric>
        );
    }

    public componentDidUpdate(previousProps: any, previousState: IDetailsListDocumentsExampleState) {
        if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
            this._selection.setAllSelected(false);
        }
    }

    private _getKey(item: any, index?: number): string {
        return item.key;
    }


    private _onRenderItemColumn(item: any, index: number, column: IColumn): JSX.Element {

        if (column.fieldName === 'Email') {
            return <OfficeUI.Link href={"mailto:" + item.Email} > {item.Email}</OfficeUI.Link>;
        }
        return item[column.fieldName];
    }

    private _getSelectionDetails(): string {
        const selectionCount = this._selection.getSelectedCount();
        this.props.UserControlAction(this._selection.getSelection());

        switch (selectionCount) {
            case 0:
                return 'No items selected';
            case 1:
                return '1 item selected: ' + (this._selection.getSelection()[0] as IDocument).Title;
            default:
                return `${selectionCount} items selected`;
        }
    }
}