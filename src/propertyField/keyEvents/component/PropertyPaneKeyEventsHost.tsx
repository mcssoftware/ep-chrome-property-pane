import * as React from "react";
import { IPropertyPaneKeyEventsHostProps, IPropertyPaneKeyEventsHostState } from "./IPropertyPaneKeyEventsHost";
import { ISPService, ISPList, ISPLists } from "../../../services/ISPService";
import { initGlobalVars } from "../../../common/ep";
import { } from "../../../common/global";
import { IPropertyFieldKeyEventsData, getKeyEventsDefaultValues } from "../IPropertyFieldKeyEvents";
import styles from "./PropertyPaneKeyEventsHost.module.scss";
import Header from "./../../header/header";
import { Spinner, SpinnerType } from "office-ui-fabric-react/lib/Spinner";
import FieldErrorMessage from "../../errorMessage/FieldErrorMessage";
import { ChoiceGroup, IChoiceGroupOption } from "office-ui-fabric-react/lib/ChoiceGroup";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Async } from "office-ui-fabric-react/lib/Utilities";
import { cloneDeep } from "@microsoft/sp-lodash-subset";

const allListSelectedKey: string = "All";
const listSelectionKey: string = "selectedList";

export default class PropertyFieldKeyEventsHost extends React.Component<IPropertyPaneKeyEventsHostProps, IPropertyPaneKeyEventsHostState> {
    private spService: ISPService;
    private allListValue: string[];
    private async: Async;
    private delayedValidate: (value: IPropertyFieldKeyEventsData) => void;

    constructor(props: IPropertyPaneKeyEventsHostProps) {
        super(props);
        if (typeof (window as any).Epmodern === "undefined") {
            initGlobalVars();
        }
        this.spService = props.spService;
        this.allListValue = [];
        this.async = new Async(this);
        this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);

        const tempValue: IPropertyFieldKeyEventsData = props.value || {} as IPropertyFieldKeyEventsData;
        const defaultValues: IPropertyFieldKeyEventsData = getKeyEventsDefaultValues();
        this.state = {
            value: {
                list: tempValue.list || defaultValues.list,
                numberOfItemsToDisplay: tempValue.numberOfItemsToDisplay || defaultValues.numberOfItemsToDisplay,
                showCalendarCenterButton: typeof tempValue.showCalendarCenterButton === "undefined" ? defaultValues.showCalendarCenterButton : tempValue.showCalendarCenterButton,
                showCalendarIcon: typeof tempValue.showCalendarIcon === "undefined" ? defaultValues.showCalendarIcon : tempValue.showCalendarIcon,
                showMonthOnTop: typeof tempValue.showMonthOnTop === "undefined" ? defaultValues.showMonthOnTop : tempValue.showMonthOnTop,
                displayStandardEvents: typeof tempValue.displayStandardEvents === "undefined" ? defaultValues.displayStandardEvents : tempValue.displayStandardEvents,
            },
            listOptions: [],
            listLoaded: false,
            listChoiceKey: listSelectionKey,
            selectedLists: tempValue.list || defaultValues.list,
            numberOfItemsText: (tempValue.numberOfItemsToDisplay || defaultValues.numberOfItemsToDisplay).toString(),
            errorMessage: ""
        };
    }

    public componentDidMount(): void {
        // Start retrieving the SharePoint lists
        this.loadLists();
    }

    public render(): JSX.Element {
        const { value, listLoaded, listChoiceKey, selectedLists, numberOfItemsText } = this.state;
        const forcedDisabled: boolean = this.props.disabled || false;
        return (
            <div className={styles.propertyPaneKeyEventsHost}>
                <Header title={this.props.label} />
                {!listLoaded && <Spinner type={SpinnerType.normal} />}
                {listLoaded && <div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <ChoiceGroup
                                selectedKey={listChoiceKey}
                                disabled={forcedDisabled}
                                className={styles.displayModeChoice}
                                options={[
                                    {
                                        key: allListSelectedKey,
                                        text: allListSelectedKey
                                    },
                                    {
                                        key: listSelectionKey,
                                        text: "Specific Calendar(s)",
                                        onRenderField: (props, render) => {
                                            return (
                                                <div>
                                                    {render!(props)}
                                                    <Dropdown
                                                        className={styles.listSelectorDdl}
                                                        multiSelect={true}
                                                        defaultSelectedKeys={selectedLists}
                                                        options={this.state.listOptions}
                                                        required={listChoiceKey === listSelectionKey}
                                                        disabled={forcedDisabled || (listChoiceKey !== listSelectionKey)}
                                                        onChanged={this.onListDropDownChanged}
                                                    />
                                                </div>
                                            );
                                        }
                                    }
                                ]}
                                onChange={this.onChoiceChanged}
                                label="Select Calendar(s)"
                            />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <TextField label="Items to display"
                                value={numberOfItemsText}
                                disabled={forcedDisabled}
                                onChanged={this.onNumberOfItemTextChanged}
                                onGetErrorMessage={this.validateNumber} />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Checkbox label="Show Calendar Icon"
                                className={styles.choiceSelector}
                                disabled={forcedDisabled}
                                checked={value.showCalendarIcon}
                                onChange={this.onShowCalendarChanged} />
                            <Checkbox label="Show Calendar Month on Top"
                                className={styles.choiceSelector}
                                disabled={forcedDisabled}
                                checked={value.showMonthOnTop}
                                onChange={this.onCalendarMonthOnTopChanged} />
                            <Checkbox label="Show Calendar Center Button"
                                className={styles.choiceSelector}
                                checked={value.showCalendarCenterButton}
                                disabled={forcedDisabled}
                                onChange={this.onShowCalendarButtonChanged} />
                            <Checkbox label="Display Standard Events"
                                className={styles.choiceSelector}
                                checked={value.displayStandardEvents}
                                disabled={forcedDisabled}
                                onChange={this.onDisplayStandardEventsChanged} />
                        </div>
                    </div>
                </div>}
                <FieldErrorMessage errorMessage={this.state.errorMessage} />
            </div>
        );
    }

    /**
     * Called when the component will unmount
     */
    public componentWillUnmount() {
        if (typeof this.async !== "undefined") {
            this.async.dispose();
        }
    }

    /**
     * Loads the list from SharePoint calendar central
     */
    private loadLists(): void {
        this.spService.getLists(this.props, window.Epmodern.urls.calendarCenterUrl).then((response: ISPLists) => {
            const options = [];
            this.allListValue = [];
            // Start mapping the list that are selected
            response.value.map((list: ISPList) => {
                options.push({
                    key: list.Id,
                    text: list.Title
                });
                this.allListValue.push(list.Id);
            });

            let { value, listChoiceKey } = this.state;
            let selectedLists: string[] = [];

            if (value.list.length === 0) {
                value.list = [...this.allListValue];
                listChoiceKey = allListSelectedKey;
            } else {
                const tempListList: string[] = [];
                for (let i: number = 0; i < value.list.length; i++) {
                    const index: number = this.allListValue.indexOf(value.list[i]);
                    if (index > -1) {
                        tempListList.push(this.allListValue[index]);
                    }
                }
                value.list = tempListList;
                if (tempListList.length === this.allListValue.length) {
                    listChoiceKey = allListSelectedKey;
                } else {
                    selectedLists = tempListList;
                }
            }

            // Update the current component state
            this.setState({
                listOptions: options,
                listLoaded: true,
                value,
                listChoiceKey,
                selectedLists,
            });
            // trigger property change
            this.delayedValidate(value);
        });
    }

    /**
     * trigger when choice field changed.
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {IChoiceGroupOption} [option]
     * @memberof PropertyFieldCalendarHost
     */
    private onChoiceChanged = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption): void => {
        if (typeof option !== "undefined") {
            const value: IPropertyFieldKeyEventsData = cloneDeep(this.state.value);
            if (option.key === allListSelectedKey) {
                value.list = [...this.allListValue];
            } else {
                value.list = [...this.state.selectedLists];
            }
            this.setState({ value, listChoiceKey: option.key });
            this.delayedValidate(value);
        }
    }

    /**
     * trigger when list dropdown changed
     *
     * @private
     * @param {IDropdownOption} option
     * @param {number} [index]
     * @memberof PropertyFieldCalendarHost
     */
    private onListDropDownChanged = (option: IDropdownOption, index?: number): void => {
        let { selectedLists } = this.state;
        const value: IPropertyFieldKeyEventsData = cloneDeep(this.state.value);
        const key: string = option.key as string;
        if (option.selected) {
            selectedLists.push(key);
        } else {
            selectedLists = selectedLists.filter(a => a !== key);
        }
        value.list = [...selectedLists];
        this.setState({ value, selectedLists });
        this.delayedValidate(value);
    }

    /**
     * Trigger when textbox for number of items is changed.
     *
     * @private
     * @param {string} newValue
     * @memberof PropertyFieldCalendarHost
     */
    private onNumberOfItemTextChanged = (newValue: string): void => {
        if (this.validateNumber(newValue).length > 0) {
            this.setState({ numberOfItemsText: newValue });
        } else {
            const value: IPropertyFieldKeyEventsData = cloneDeep(this.state.value);
            value.numberOfItemsToDisplay = parseInt(newValue);
            this.setState({ value, numberOfItemsText: newValue });
            this.delayedValidate(value);
        }
    }

    /**
     * Trigger when show calendar checkbox is clicked
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {boolean} [checked]
     * @memberof PropertyFieldCalendarHost
     */
    private onShowCalendarChanged = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        const value: IPropertyFieldKeyEventsData = cloneDeep(this.state.value);
        value.showCalendarIcon = checked;
        this.setState({ value });
        this.delayedValidate(value);
    }

    /**
     * Trigger when calendar month on top checkbox is clicked
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {boolean} [checked]
     * @memberof PropertyFieldCalendarHost
     */
    private onCalendarMonthOnTopChanged = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        const value: IPropertyFieldKeyEventsData = cloneDeep(this.state.value);
        value.showMonthOnTop = checked;
        this.setState({ value });
        this.delayedValidate(value);
    }

    /**
     * Trigger when show calendar checkbox is clicked
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {boolean} [checked]
     * @memberof PropertyFieldCalendarHost
     */
    private onShowCalendarButtonChanged = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        const value: IPropertyFieldKeyEventsData = cloneDeep(this.state.value);
        value.showCalendarCenterButton = checked;
        this.setState({ value });
        this.delayedValidate(value);
    }

    /**
     * Trigger when display standard events checkbox is clicked
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {boolean} [checked]
     * @memberof PropertyFieldCalendarHost
     */
    private onDisplayStandardEventsChanged = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        const value: IPropertyFieldKeyEventsData = cloneDeep(this.state.value);
        value.displayStandardEvents = checked;
        this.setState({ value });
        this.delayedValidate(value);
    }

    /**
     * Validates the new custom field value
     */
    private validate = (value: IPropertyFieldKeyEventsData): void => {
        const internalResult: string = this.validateInternal(value);
        if (internalResult.length < 1) {
            if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
                this.setState({ errorMessage: "" });
                this.notifyAfterValidate(this.props.value, value);
                return;
            }
        }
        const result: string | PromiseLike<string> = internalResult.length > 0 ? internalResult : this.props.onGetErrorMessage(value || getKeyEventsDefaultValues());
        if (typeof result !== "undefined") {
            if (typeof result === "string") {
                if (result === "") {
                    this.notifyAfterValidate(this.props.value, value);
                }
                this.setState({
                    errorMessage: result
                });
            } else {
                result.then((errorMessage: string) => {
                    if (typeof errorMessage === "undefined" || errorMessage === "") {
                        this.notifyAfterValidate(this.props.value, value);
                    }
                    this.setState({
                        errorMessage: errorMessage
                    });
                });
            }
        } else {
            this.setState({ errorMessage: "" });
            this.notifyAfterValidate(this.props.value, value);
        }
    }

    private validateInternal = (value: IPropertyFieldKeyEventsData): string => {
        if (value.list.length < 1) {
            return "Please select at least one calendar list.";
        }
        if (value.numberOfItemsToDisplay < 1) {
            return "Invalid number of items to display.";
        }
        return "";
    }

    private validateNumber = (value: string): string => {
        var r = RegExp(/(^[^\-]{0,1})?(^[\d]*)$/);
        if (r.test(value) && value.length > 0 && value !== "0") {
            return "";
        }
        return "The value should be a number greater than 0.";
    }

    /**
    * Notifies the parent Web Part of a property value change
    */
    private notifyAfterValidate = (oldValue: IPropertyFieldKeyEventsData, newValue: IPropertyFieldKeyEventsData) => {
        if (this.props.onPropertyChange && newValue !== null) {
            // this.props.properties[this.props.targetProperty] = newValue;
            this.props.onPropertyChange(this.props.targetProperty, oldValue, newValue);
            // Trigger the apply button
            if (typeof this.props.onChange !== "undefined" && this.props.onChange !== null) {
                this.props.onChange(this.props.targetProperty, newValue);
            }
        }
    }

}