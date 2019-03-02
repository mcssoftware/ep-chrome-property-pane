import * as React from "react";
import { IContentData, getContentDataDefaultValue } from "../IPropertyPaneMultiZoneSelector";
import styles from "./PropertyFieldMultiZoneSelectorHost.module.scss";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { cloneDeep, isEqual } from "@microsoft/sp-lodash-subset";
import { ColorPicker } from "office-ui-fabric-react/lib/ColorPicker";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { PrimaryButton, DefaultButton } from "office-ui-fabric-react/lib/Button";

export interface IContentControlProps {
    data: IContentData;
    /**
     * Defines a onPropertyChange function to raise when the selected value changed.
     * Normally this function must be always defined with the 'this.onPropertyChange'
     * method of the web part object.
     */
    notify(oldValue: IContentData, newValue: IContentData): void;
    /**
     * Whether the property pane field is enabled or not.
     */
    disabled?: boolean;
}

export interface IContentControlState {
    value: IContentData;
    useImage: boolean;
    showPanel: boolean;
    bgColor: string;
}

/**
 * Control used for content type zone.
 * @export
 * @class ContentControl
 * @extends {React.Component<IContentControlProps, IContentControlState>}
 */
export class ContentControl extends React.Component<IContentControlProps, IContentControlState> {
    /**
     * Content control constructor
     */
    constructor(props: IContentControlProps) {
        super(props);
        const value = props.data || getContentDataDefaultValue();
        let useImage: boolean = false;
        if (typeof value.iconUrl === "string" && value.iconUrl.length > 0) {
            useImage = true;
        }
        this.state = {
            value,
            useImage,
            showPanel: false,
            bgColor: value.backgroundColor,
        };
    }

    /**
     * Render content control
     * @returns {JSX.Element}
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
     */
    public render(): JSX.Element {
        const { value, useImage } = this.state;
        const forcedDisabled: boolean = this.props.disabled || false;
        return (
            <div className={styles.propertyFieldMultiZoneNewsSelectorHost}>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <TextField label="Title"
                            value={value.title}
                            disabled={forcedDisabled}
                            onChanged={this.onTitleTextChanged} />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <Toggle label="Background Color"
                            onText="Use Image as background"
                            offText="Use color as background"
                            onChanged={this.onBackgroundToggled}
                            disabled={forcedDisabled}
                            checked={useImage} />
                    </div>
                    <div className={styles.column}>
                        {!useImage &&
                            <TextField label="TextField with an icon"
                                value={value.backgroundColor}
                                readOnly={true}
                                onClick={this.openBgColorPanel}
                                disabled={forcedDisabled}
                                iconProps={{ iconName: "Tag" }} />
                        }
                        {useImage && <TextField label="Image Url"
                            value={value.backgroundUrl}
                            onChanged={this.onBgImageUrlTextChanged} />}
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <TextField label="Icon Url"
                            value={value.iconUrl}
                            onChanged={this.onIconUrlTextChanged} />
                    </div>
                </div>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <TextField label="Target Url"
                            value={value.targetUrl}
                            disabled={forcedDisabled}
                            onChanged={this.ontargetUrlTextChanged} />
                    </div>
                </div>
                <Panel
                    isOpen={this.state.showPanel}
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

    /**
     * Static method which is invoked after a component is instantiated as well as when it receives new props
     * @static
     * @param {IContentControlProps} nextProps
     * @param {IContentControlProps} prevState
     * @returns {*}
     * @memberof ContentControl
     */
    public static getDerivedStateFromProps(nextProps: IContentControlProps, prevState: IContentControlState): any {
        if (!isEqual(nextProps.data, prevState.value)) {
            return {
                value: nextProps.data,
                useImage: typeof nextProps.data !== "undefined"
                    && typeof nextProps.data.backgroundUrl === "string" && nextProps.data.backgroundUrl.length > 0 ? true : false,
                showPanel: false,
                bgColor: typeof nextProps.data.backgroundUrl === "string" ? "#ffffff" : nextProps.data.backgroundColor
            };
        }
        return null;
    }

    /**
     *
     *
     * @private
     * @memberof ContentControl
     */
    private openBgColorPanel = (): void => {
        if (this.props.disabled !== true) {
            this.setState({ showPanel: true, bgColor: this.state.value.backgroundColor });
        }
    }

    /**
     *
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
     *
     *
     * @private
     * @memberof ContentControl
     */
    private onPanelColorSaved = (): void => {
        const value = cloneDeep(this.state.value);
        value.backgroundColor = this.state.bgColor;
        this.setState({ showPanel: false, value });
        this.validate(value);
    }

    /**
     * On backgroud color choice changed
     * @private
     * @memberof ContentControl
     */
    private onBackgroundToggled = (checked: boolean): void => {
        this.setState({ useImage: checked });
    }

    /**
     * On color picker panel closed.
     *
     * @private
     * @memberof ContentControl
     */
    private onColorPickerPanelClose = (): void => {
        this.setState({ showPanel: false });
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
     * On background image url changed.
     *
     * @private
     * @memberof ContentControl
     */
    private onBgImageUrlTextChanged = (textValue: string): void => {
        const value = cloneDeep(this.state.value);
        value.backgroundUrl = textValue;
        this.setState({ value });
        this.validate(value);
    }

    /**
     * On icon url changed.
     *
     * @private
     * @memberof ContentControl
     */
    private onIconUrlTextChanged = (textValue: string): void => {
        const value = cloneDeep(this.state.value);
        value.iconUrl = textValue;
        this.setState({ value });
        this.validate(value);
    }

    /**
     * On target url changed.
     *
     * @private
     * @memberof ContentControl
     */
    private ontargetUrlTextChanged = (textValue: string): void => {
        const value = cloneDeep(this.state.value);
        value.targetUrl = textValue;
        this.setState({ value });
        this.validate(value);
    }

    /**
     * On title text changed
     *
     * @private
     * @param {string} textValue
     * @memberof ContentControl
     */
    private onTitleTextChanged = (textValue: string): void => {
        const value = cloneDeep(this.state.value);
        value.title = textValue;
        this.setState({ value });
        this.validate(value);
    }

    /**
     * Validates the new custom field value
     * @private
     * @param {IPropertyPaneMultiZoneSelectorData} value
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
    */
    private validate = (newValue: IContentData): void => {
        const internalResult: string = this.validateInternal(newValue);
        if (internalResult.length === 0) {
            this.notifyAfterValidate(newValue);
        }
    }

    private validateInternal = (newValue: IContentData): string => {
        return "";
    }

    private notifyAfterValidate = (newValue: IContentData) => {
        if (typeof this.props.notify === "function") {
            this.props.notify(this.props.data, newValue);
        }
    }
}