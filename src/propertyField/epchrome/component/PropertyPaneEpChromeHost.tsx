import * as React from "react";
import { IPropertyFieldEpChromeData, getEpChromeDataDefaultValues } from "../IPropertyFieldEpChrome";
import { IPropertyPaneEpChromeHostProps, IPropertyPaneEpChromeHostState } from "./IPropertyPaneEpChromeHost";
import styles from "./PropertyPaneEpChromeHost.module.scss";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import { TextField } from "office-ui-fabric-react";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import Header from '../../header/header';

export default class PropertPaneEpChromeHost extends React.Component<IPropertyPaneEpChromeHostProps, IPropertyPaneEpChromeHostState> {

    constructor(props: IPropertyPaneEpChromeHostProps) {
        super(props);
        const tempValue: IPropertyFieldEpChromeData = props.value || {} as IPropertyFieldEpChromeData;
        const defaultValues = getEpChromeDataDefaultValues();
        this.validate = this.validate.bind(this);
        this.state = {
            value: {
                IconPath: tempValue.IconPath || defaultValues.IconPath,
                ShowTitle: tempValue.ShowTitle || defaultValues.ShowTitle,
                Title: tempValue.Title || defaultValues.Title
            }
        };
    }

    public render(): JSX.Element {
        const { value } = this.state;
        const label: string = this.props.label || "Ep Chrome Settings";
        return (
            <div className={styles.propertyPaneEpChromeHost}>
                <Header title={label} />
                <div className={styles.row}>
                    <div className={styles.column}>
                        <Toggle checked={value.ShowTitle}
                            label="Show Chrome"
                            onChanged={this._onShowTitleToggleChanged}
                        />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <TextField label="Title"
                            value={value.Title}
                            disabled={!value.ShowTitle}
                            required={value.ShowTitle}
                            onChanged={this._onTitleChanged}
                            errorMessage={this._getTitleErrorMessage()} />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <TextField label="Icon Url"
                            value={value.IconPath}
                            onChanged={this._onIconUrlChanged}
                        />
                    </div>
                </div>
            </div>
        );
    }

    private _onShowTitleToggleChanged = (checked: boolean): void => {
        const value = cloneDeep(this.state.value);
        value.ShowTitle = checked;
        this.setState({ value });
        this.validate(value);
    }

    private _onTitleChanged = (textValue: string): void => {
        const value: IPropertyFieldEpChromeData = cloneDeep(this.state.value);
        value.Title = textValue;
        this.setState({ value });
        this.validate(value);
    }

    private _onIconUrlChanged = (textValue: string): void => {
        const value: IPropertyFieldEpChromeData = cloneDeep(this.state.value);
        value.IconPath = textValue;
        this.setState({ value });
        this.validate(value);
    }

    private _getTitleErrorMessage = (): string => {
        const { value } = this.state;
        if (value.ShowTitle && value.Title.trim().length === 0) {
            return "Title is required";
        }
        return "";
    }

   /**
   * Validates the new custom field value
   */
    private validate(value: IPropertyFieldEpChromeData): void {
        if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
            this.notifyAfterValidate(this.props.value, value);
            return;
        }
        const result: string | PromiseLike<string> = this.props.onGetErrorMessage(value);
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
            this.notifyAfterValidate(this.props.value, value);
        }
    }

    /**
   * Notifies the parent Web Part of a property value change
   */
    private notifyAfterValidate(oldValue: IPropertyFieldEpChromeData, newValue: IPropertyFieldEpChromeData) {
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
