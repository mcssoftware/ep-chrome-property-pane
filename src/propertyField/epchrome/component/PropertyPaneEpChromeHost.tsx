import * as React from "react";
import { IPropertyFieldEpChromeData, getEpChromeDataDefaultValues } from "../IPropertyFieldEpChrome";
import { IPropertyPaneEpChromeHostProps, IPropertyPaneEpChromeHostState } from "./IPropertyPaneEpChromeHost";
import styles from "./PropertyPaneEpChromeHost.module.scss";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import Header from '../../header/header';
import { PanelType, Panel } from "office-ui-fabric-react/lib/Panel";
import { ColorPicker } from "office-ui-fabric-react/lib/ColorPicker";
import { PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/components/Button";

export default class PropertPaneEpChromeHost extends React.Component<IPropertyPaneEpChromeHostProps, IPropertyPaneEpChromeHostState> {

    constructor(props: IPropertyPaneEpChromeHostProps) {
        super(props);
        const tempValue: IPropertyFieldEpChromeData = props.value || {} as IPropertyFieldEpChromeData;
        const defaultValues = getEpChromeDataDefaultValues();
        this.validate = this.validate.bind(this);
        this.state = {
            value: {
                IconPath: tempValue.IconPath || defaultValues.IconPath,
                ShowTitle: typeof tempValue.ShowTitle === "undefined" ? defaultValues.ShowTitle : tempValue.ShowTitle,
                Title: tempValue.Title || defaultValues.Title,
                BackgroundColor: tempValue.BackgroundColor || defaultValues.BackgroundColor,
            },
            showColorPanel: false,
            errorMessage: "",
            bgColor: tempValue.BackgroundColor || defaultValues.BackgroundColor,
        };
    }

    public render(): JSX.Element {
        const { value } = this.state;
        const label: string = this.props.label;
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
                <div className={styles.row}>
                    <div className={styles.column}>
                        <TextField label="Background Color"
                            value={value.BackgroundColor}
                            readOnly={true}
                            onClick={this.openBgColorPanel}
                            iconProps={{ iconName: "Tag" }} />
                    </div>
                </div>
                <Panel
                    isOpen={this.state.showColorPanel}
                    type={PanelType.smallFixedFar}
                    onDismiss={this.onColorPickerPanelClose}
                    headerText="Color Picker"
                    closeButtonAriaLabel="Close"
                    onRenderFooterContent={this.onPanelRenderFooterContent}>
                    <div>
                        <ColorPicker color={this.state.bgColor}
                            onColorChanged={this.onColorChanged}
                            alphaSliderHidden={true} />
                    </div>
                    <div className={styles.column}>
                        <div style={{ width: "100px", height: "100px", backgroundColor: this.state.bgColor, borderStyle: "solid", borderWidth: "1px", borderColor: "#cdcdcd" }}>
                        </div>
                    </div>
                </Panel>
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
     * Adding buttons on footer of panel
     *
     * @private
     * @returns {JSX.Element}
     * @memberof ContentControl
     */
    private onPanelRenderFooterContent = (): JSX.Element => {
        return (
            <div>
                <PrimaryButton onClick={this.onPanelColorSaved} style={{ marginRight: '8px' }}>
                    Save
                </PrimaryButton>
                <DefaultButton onClick={this.onColorPickerPanelClose}>Cancel</DefaultButton>
            </div>
        );
    }

    /**
     * On color picker panel closed.
     *
     * @private
     * @memberof ContentControl
     */
    private onColorPickerPanelClose = (): void => {
        this.setState({ showColorPanel: false });
    }

    /**
     * Color selected.
     * @private
     * @memberof ContentControl
     */
    private onColorChanged = (color: string): void => {
        this.setState({ bgColor: color });
    }

    /**
     *
     *
     * @private
     * @memberof ContentControl
     */
    private onPanelColorSaved = (): void => {
        const value = cloneDeep(this.state.value);
        value.BackgroundColor = this.state.bgColor;
        this.setState({ showColorPanel: false, value });
        this.validate(value);
    }

    /**
     *
     *
     * @private
     * @memberof ContentControl
     */
    private openBgColorPanel = (): void => {
        this.setState({ showColorPanel: true, bgColor: this.state.value.BackgroundColor });
    }

    /**
    * Validates the new custom field value
    */
    private validate = (value: IPropertyFieldEpChromeData): void => {
        if (this.props.onGetErrorMessage === null || this.props.onGetErrorMessage === undefined) {
            this.setState({ errorMessage: "" });
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
            this.setState({ errorMessage: "" });
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
