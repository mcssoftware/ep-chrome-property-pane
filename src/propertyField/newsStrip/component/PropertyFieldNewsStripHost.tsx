import * as React from "react";
import { IPropertyFieldNewsStripHostProps, IPropertyFieldNewsStripHostHost } from "./IPropertyFieldNewsStripHost";
import { IPropertyFieldNewsStripData, getNewsStripDefaultValues } from "../IPropertyFieldNewsStrip";
import { Async } from "office-ui-fabric-react/lib/Utilities";
import styles from "./PropertyFieldNewsStripHost.module.scss";
import Header from "../../header/header";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";

export default class PropertyFieldNewsStripHost extends React.Component<IPropertyFieldNewsStripHostProps, IPropertyFieldNewsStripHostHost>{
    private async: Async;
    private delayedValidate: (value: IPropertyFieldNewsStripData) => void;

    constructor(props: IPropertyFieldNewsStripHostProps) {
        super(props);
        this.async = new Async(this);
        this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
        const tempValue: IPropertyFieldNewsStripData = props.value || {} as IPropertyFieldNewsStripData;
        const defaultValues = getNewsStripDefaultValues();
        const value: IPropertyFieldNewsStripData = {
            numberOfItemsToDisplay: tempValue.numberOfItemsToDisplay || defaultValues.numberOfItemsToDisplay,
            showAuthor: typeof tempValue.showAuthor === "undefined" ? defaultValues.showAuthor : tempValue.showAuthor,
            showArticleDate: typeof tempValue.showArticleDate === "undefined" ? defaultValues.showArticleDate : tempValue.showArticleDate,
            showRating: typeof tempValue.showRating === "undefined" ? defaultValues.showRating : tempValue.showRating,
            showSummary: typeof tempValue.showSummary === "undefined" ? defaultValues.showSummary : tempValue.showSummary,
        };
        this.state = {
            value,
            numberOfItemsText: value.numberOfItemsToDisplay.toString(),
        };
    }

    /**
     * Called when the component will unmount
     * @memberof PropertyFieldNewsStripHost
     */
    public componentWillUnmount() {
        if (typeof this.async !== "undefined") {
            this.async.dispose();
        }
    }

    public render(): JSX.Element {
        const { value, numberOfItemsText } = this.state;
        const forcedDisabled: boolean = this.props.disabled || false;
        return (
            <div className={styles.propertyPaneNewsStripHost}>
                <Header title={this.props.label} />
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
                        <Checkbox label="Show Author "
                            className={styles.choiceSelector}
                            disabled={forcedDisabled}
                            checked={value.showAuthor}
                            onChange={this.onShowAuthorChanged} />
                        <Checkbox label="Show Article Date"
                            className={styles.choiceSelector}
                            disabled={forcedDisabled}
                            checked={value.showArticleDate}
                            onChange={this.onShowArticleDateChanged} />
                        <Checkbox label="Show Ratings"
                            className={styles.choiceSelector}
                            disabled={forcedDisabled}
                            checked={value.showRating}
                            onChange={this.onShowRatingChanged} />
                        <Checkbox label="Show Summary"
                            className={styles.choiceSelector}
                            disabled={forcedDisabled}
                            checked={value.showSummary}
                            onChange={this.onShowSummaryChanged} />
                    </div>
                </div>
            </div>
        );
    }

    /**
     * Trigger when textbox for number of items is changed.
     *
     * @private
     * @param {string} newValue
     * @memberof PropertyFieldNewsStripHost
     */
    private onNumberOfItemTextChanged = (newValue: string): void => {
        if (this.validateNumber(newValue).length > 0) {
            this.setState({ numberOfItemsText: newValue });
        } else {
            const value: IPropertyFieldNewsStripData = cloneDeep(this.state.value);
            value.numberOfItemsToDisplay = parseInt(newValue);
            this.setState({ value, numberOfItemsText: newValue });
            this.delayedValidate(value);
        }
    }

    /**
     * Trigger when show article date checkbox is clicked
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {boolean} [checked]
     * @memberof PropertyFieldNewsStripHost
     */
    private onShowArticleDateChanged = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        const value: IPropertyFieldNewsStripData = cloneDeep(this.state.value);
        value.showArticleDate = checked;
        this.setState({ value });
        this.delayedValidate(value);
    }

    /**
     * Trigger when show author checkbox is clicked
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {boolean} [checked]
     * @memberof PropertyFieldNewsStripHost
     */
    private onShowAuthorChanged = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        const value: IPropertyFieldNewsStripData = cloneDeep(this.state.value);
        value.showAuthor = checked;
        this.setState({ value });
        this.delayedValidate(value);
    }

    /**
     * Trigger when show rating checkbox is clicked
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {boolean} [checked]
     * @memberof PropertyFieldNewsStripHost
     */
    private onShowRatingChanged = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        const value: IPropertyFieldNewsStripData = cloneDeep(this.state.value);
        value.showRating = checked;
        this.setState({ value });
        this.delayedValidate(value);
    }

    /**
     * Trigger when show summary checkbox is clicked
     *
     * @private
     * @param {(React.FormEvent<HTMLElement | HTMLInputElement>)} [ev]
     * @param {boolean} [checked]
     * @memberof PropertyFieldNewsStripHost
     */
    private onShowSummaryChanged = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean): void => {
        const value: IPropertyFieldNewsStripData = cloneDeep(this.state.value);
        value.showSummary = checked;
        this.setState({ value });
        this.delayedValidate(value);
    }

    /**
     * Validates the new custom field value
     *
     * @private
     * @param {IPropertyFieldNewsStripData} value
     * @returns {void}
     * @memberof PropertyFieldNewsStripHost
     */
    private validate = (value: IPropertyFieldNewsStripData): void => {
        const internalResult: string = this.validateInternal(value);
        if (internalResult.length < 1) {
            if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
                this.setState({ errorMessage: "" });
                this.notifyAfterValidate(this.props.value, value);
                return;
            }
        }
        const result: string | PromiseLike<string> = internalResult.length > 0 ? internalResult : this.props.onGetErrorMessage(value || getNewsStripDefaultValues());
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

    private validateInternal(value: IPropertyFieldNewsStripData): string {
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
     *
     * @private
     * @param {IPropertyFieldNewsStripData} oldValue
     * @param {IPropertyFieldNewsStripData} newValue
     * @memberof PropertyFieldNewsStripHost
     */
    private notifyAfterValidate(oldValue: IPropertyFieldNewsStripData, newValue: IPropertyFieldNewsStripData) {
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