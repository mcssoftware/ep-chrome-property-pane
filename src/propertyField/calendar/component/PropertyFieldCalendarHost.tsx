import * as React from "react";
import { IPropertyPaneCalendarHostProps, IPropertyPaneCalendarHostState } from "./IPropertyFieldCalendarHost";
import { IPropertyFieldCalendarData, getCalendarDataDefaultValues, CalendarDisplayModeType } from "../IPropertyFieldCalendar";
import styles from "./PropertyFieldCalendarHost.module.scss";
import Header from "./../../header/header";
import FieldErrorMessage from "../../errorMessage/FieldErrorMessage";
import { Async } from "office-ui-fabric-react/lib/Utilities";
import { IDropdownOption, Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { ISPService, ISPLists, ISPList } from "../../../services/ISPService";
import { initGlobalVars } from "../../../common/ep";
import { Spinner, SpinnerType } from "office-ui-fabric-react/lib/Spinner";
import { ChoiceGroup, IChoiceGroupOption } from "office-ui-fabric-react/lib/ChoiceGroup";
import { cloneDeep } from "@microsoft/sp-lodash-subset";

export default class PropertyFieldCalendarHost extends React.Component<IPropertyPaneCalendarHostProps, IPropertyPaneCalendarHostState> {
    private async: Async;
    private delayedValidate: (value: IPropertyFieldCalendarData) => void;
    private spService: ISPService;

    constructor(props: IPropertyPaneCalendarHostProps) {
        super(props);
        if (typeof window.Epmodern === "undefined") {
            initGlobalVars();
        }
        this.spService = props.spService;
        const tempValue: IPropertyFieldCalendarData = props.value || {} as IPropertyFieldCalendarData;
        const defaultValues = getCalendarDataDefaultValues();
        this.validate = this.validate.bind(this);
        this.onListSelectionDdlChanged = this.onListSelectionDdlChanged.bind(this);
        this.onDisplayModeChange = this.onDisplayModeChange.bind(this);
        this.onItemDropDownChanged = this.onItemDropDownChanged.bind(this);
        this.onRenderDropDownTitle = this.onRenderDropDownTitle.bind(this);
        this.async = new Async(this);
        this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
        const value = {
            CalendarDisplayMode: tempValue.CalendarDisplayMode || defaultValues.CalendarDisplayMode,
            CalendarId: tempValue.CalendarId || defaultValues.CalendarId,
            ListId: tempValue.ListId || defaultValues.ListId,
            ListTitle: tempValue.ListTitle || defaultValues.ListTitle
        };
        this.state = {
            value,
            listOptions: [],
            listLoaded: false,
            itemsDropDownOptions: [],
            itemsLoaded: true,
            errorMessage: this.validateInternal(value)
        };
    }

    public componentDidMount(): void {
        // Start retrieving the SharePoint lists
        this.loadLists();
        this.loadListItems(this.state.value);
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
            // Start mapping the list that are selected
            response.value.map((list: ISPList) => {
                options.push({
                    key: list.Id,
                    text: list.Title
                });
            });

            // Option to unselect the list
            options.unshift({
                key: "",
                text: "Select List"
            });

            // Update the current component state
            this.setState({
                listOptions: options,
                listLoaded: true,
            });
        });
    }

    /**
     * Load List items from Sharepoint List
     *
     * @private
     * @param {IPropertyFieldCalendarData} value
     * @memberof PropertyFieldCalendarHost
     */
    private loadListItems(value: IPropertyFieldCalendarData): void {
        if (value.ListId.length > 0) {
            this.setState({ itemsLoaded: false, itemsDropDownOptions: this.getEmptyDropDownOption() });
            const now: string = new Date(Date.now()).toISOString();
            const filter: string = `EventDate ge '${now}'`;
            this.spService.getListItems(filter, value.ListTitle, "Title", window.Epmodern.urls.calendarCenterUrl, "Title", 100)
                .then((items) => {
                    const itemsDropDownOptions: IDropdownOption[] = this.getEmptyDropDownOption();
                    items.filter((f) => typeof f.Title === "string" && f.Title.trim().length > 0).forEach(f => {
                        itemsDropDownOptions.push({
                            key: f.Id.toString(),
                            text: f.Title.trim()
                        });
                    });
                    this.setState({ itemsLoaded: true, itemsDropDownOptions });
                });
        } else {
            this.setState({ itemsDropDownOptions: this.getEmptyDropDownOption() });
        }
    }

    /**
     * Empty options for list items dropdown
     *
     */
    private getEmptyDropDownOption(): IDropdownOption[] {
        return [{ key: "0", text: "Select Item" }];
    }

    public render(): JSX.Element {
        const { value, listLoaded } = this.state;
        const forcedDisabled: boolean = this.props.disabled || false;
        return (
            <div className={styles.propertyPaneCalendarHost}>
                <Header title={this.props.label} />
                {!listLoaded && <Spinner type={SpinnerType.normal} />}
                {listLoaded && <div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <Dropdown
                                disabled={forcedDisabled}
                                label="Select Calendar"
                                onChanged={this.onListSelectionDdlChanged}
                                options={this.state.listOptions}
                                selectedKey={this.state.value.ListId}
                            />
                        </div>
                    </div>
                    <div className={styles.row}>
                        <div className={styles.column}>
                            <ChoiceGroup
                                selectedKey={value.CalendarDisplayMode.toString()}
                                disabled={forcedDisabled}
                                className={styles.displayModeChoice}
                                options={[
                                    {
                                        key: CalendarDisplayModeType.Latest.toString(),
                                        text: "Nearest upcoming event"
                                    },
                                    {
                                        key: CalendarDisplayModeType.Specific.toString(),
                                        text: "Specific event",
                                        onRenderField: (props, render) => {
                                            return (
                                                <div>
                                                    {render!(props)}
                                                    <Dropdown
                                                        className={styles.pageSelectorDdl}
                                                        selectedKey={value.CalendarId.toString()}
                                                        onRenderTitle={this.onRenderDropDownTitle}
                                                        options={this.state.itemsDropDownOptions}
                                                        required={value.CalendarDisplayMode === CalendarDisplayModeType.Specific}
                                                        disabled={forcedDisabled || (value.CalendarDisplayMode === CalendarDisplayModeType.Latest)}
                                                        onChanged={this.onItemDropDownChanged}
                                                    />
                                                </div>
                                            );
                                        }
                                    }
                                ]}
                                onChange={this.onDisplayModeChange}
                                label="Calendar Display Mode"
                                required={true}
                            />
                        </div>
                    </div>
                </div>}
                <FieldErrorMessage errorMessage={this.state.errorMessage} />
            </div>
        );
    }

    /**
   * Raises when a list has been selected
   */
    private onListSelectionDdlChanged(option: IDropdownOption, index?: number): void {
        const data = cloneDeep(this.state.value);
        data.ListId = option.key as string;
        if (data.ListId.length > 0) {
            data.ListTitle = option.text;
        }
        this.setState({ value: data });
        this.delayedValidate(data);
        this.loadListItems(data);
    }

    /**
     * Raises when calendar display more is selected or changed
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {IChoiceGroupOption} [option]
     * @memberof PropertyFieldCalendarHost
     */
    private onDisplayModeChange(ev?: React.FormEvent<HTMLElement | HTMLInputElement>, option?: IChoiceGroupOption): void {
        if (typeof option !== "undefined") {
            let mode: CalendarDisplayModeType = CalendarDisplayModeType.Latest;
            if (option.key === CalendarDisplayModeType.Specific.toString()) {
                mode = CalendarDisplayModeType.Specific;
            }
            const data = cloneDeep(this.state.value);
            data.CalendarDisplayMode = mode;
            this.setState({ value: data });
            this.delayedValidate(data);
        }
    }

    /**
     * Raises when items are selected
     *
     * @private
     * @param {IDropdownOption} option
     * @param {number} [index]
     * @memberof PropertyFieldCalendarHost
     */
    private onItemDropDownChanged(option: IDropdownOption, index?: number): void {
        const key: string = option.key as string;
        const data = cloneDeep(this.state.value);
        data.CalendarId = parseInt(key);
        this.setState({ value: data });
        this.delayedValidate(data);
    }

    /**
     * 
     *
     * @private
     * @param {IDropdownOption[]} options
     * @returns {JSX.Element}
     * @memberof PropertyFieldCalendarHost
     */
    private onRenderDropDownTitle(options: IDropdownOption[]): JSX.Element {
        const option = options[0];
        return (
            <div>
                {!this.state.itemsLoaded && <Spinner type={SpinnerType.normal} />}
                {this.state.itemsLoaded && <span>{option.text}</span>}
            </div>
        );
    }

    /**
   * Validates the new custom field value
   */
    private validate = (value: IPropertyFieldCalendarData): void => {
        const internalResult: string = this.validateInternal(value);
        if (internalResult.length < 1) {
            if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
                this.setState({ errorMessage: "" });
                this.notifyAfterValidate(this.props.value, value);
                return;
            }
        }
        const result: string | PromiseLike<string> = internalResult.length > 0 ? internalResult : this.props.onGetErrorMessage(value || getCalendarDataDefaultValues());
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

    private validateInternal(value: IPropertyFieldCalendarData): string {
        if (typeof value === "undefined" || value === null) {
            return "";
        }
        if (typeof value.ListId !== "string" || value.ListId === "") {
            return "Please select list";
        }
        if (typeof value.ListTitle !== "string" || value.ListTitle === "") {
            return "Please select list";
        }
        const articleIdValid: boolean = typeof value.CalendarId !== "undefined" && value.CalendarId !== null && !isNaN(parseFloat(value.CalendarId.toString())) && isFinite(value.CalendarId);
        if ((value.CalendarDisplayMode === CalendarDisplayModeType.Specific) && !articleIdValid) {
            return "Please select calendar item";
        }
        return "";
    }

    /**
    * Notifies the parent Web Part of a property value change
    */
    private notifyAfterValidate(oldValue: IPropertyFieldCalendarData, newValue: IPropertyFieldCalendarData) {
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