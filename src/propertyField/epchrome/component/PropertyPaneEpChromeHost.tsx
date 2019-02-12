import * as React from 'react';
import { IPropertyFieldEpChromeData } from "../IPropertyFieldEpChrome";
import { IPropertyPaneEpChromeHostProps, IPropertyPaneEpChromeHostState } from "./IPropertyPaneEpChromeHost";
import styles from "./PropertyPaneEpChromeHost.module.scss";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import { TextField, Dropdown, IDropdownOption, Icon } from "office-ui-fabric-react";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import { FabricIconNames } from "./FabricIconNames";
import { initializeIcons } from '@uifabric/icons';
initializeIcons();

export default class PropertPaneEpChromeHost extends React.Component<IPropertyPaneEpChromeHostProps, IPropertyPaneEpChromeHostState> {
    private _iconOptions: any[];

    constructor(props: IPropertyPaneEpChromeHostProps) {
        super(props);
        const tempValue = props.value || {};
        this._initializeIconOptions();
        this.state = {
            value: {
                ...tempValue, ...this._getDefaultValues()
            }
        };
    }

    public render(): JSX.Element {
        const { value } = this.state;
        return (
            <div className={styles.propertyPaneEpChromeHost}>
                <Toggle checked={value.showTitle}
                    label="Show Chrome"
                    onChanged={this._onShowTitleToggleChanged}
                />
                <TextField label="Title"
                    value={value.title}
                    disabled={!value.showTitle}
                    required={value.showTitle}
                    onChanged={this._onTitleChanged}
                    errorMessage={this._getTitleErrorMessage()} />
                {/* <Toggle checked={value.showIcon}
                    label="Show Icon"
                    onChanged={this._onShowIconToggleChanged}
                /> */}
                <Dropdown
                    placeholder="Select an Icon"
                    label="Select an Icon"
                    // disabled={!value.showIcon}
                    // required={value.showIcon}
                    onRenderTitle={this._onRenderTitle}
                    onRenderOption={this._onRenderOption}
                    options={this._iconOptions}
                    errorMessage={this._getIconErrorMessage()}
                />
            </div>
        );
    }

    private _getDefaultValues(): IPropertyFieldEpChromeData {
        return {
            iconPath: "",
            isActive: false,
            showIcon: false,
            showTitle: false,
            title: "",
        };
    }

    private _onShowTitleToggleChanged = (checked: boolean): void => {
        const value = cloneDeep(this.state.value);
        value.showTitle = checked;
        this.setState({ value });
    }

    private _onShowIconToggleChanged = (checked: boolean): void => {
        const value = cloneDeep(this.state.value);
        value.showIcon = checked;
        this.setState({ value });
    }

    private _onRenderOption = (option: IDropdownOption): JSX.Element => {
        return (
            <div className="dropdownExample-option">
                {(option.key as string).length > 0 && (
                    <Icon style={{ marginRight: '8px' }} iconName={option.text} aria-hidden="true" title={option.text} />
                )}
                <span>{option.text}</span>
            </div>
        );
    }

    private _onRenderTitle = (options: IDropdownOption[]): JSX.Element => {
        const option = options[0];

        return (
            <div className="dropdownExample-option">
                {(option.key as string).length > 0 && (
                    <Icon style={{ marginRight: '8px' }} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
                )}
                <span>{option.text}</span>
            </div>
        );
    }

    private _onTitleChanged = (textValue: string): void => {
        const value: IPropertyFieldEpChromeData = cloneDeep(this.state.value);
        value.title = textValue;
        this.setState({ value });
    }

    private _getTitleErrorMessage = (): string => {
        const { value } = this.state;
        if (value.showTitle && value.title.trim().length === 0) {
            return "Title is required";
        }
        return "";
    }

    private _getIconErrorMessage = (): string => {
        const { value } = this.state;
        if (value.showTitle && value.title.trim().length === 0) {
            return "Icon is required";
        }
        return "";
    }

    private _initializeIconOptions = (): void => {
        this._iconOptions = [{ key: "", text: "Select an Icon" }];
        for (var prop in FabricIconNames) {
            const iconname: string = FabricIconNames[prop];
            this._iconOptions.push({ key: iconname, text: iconname });
        }
    }

    /**
   * Validate if field value is a number
   * @param value
   */
    private _validateNumber = (value: IPropertyFieldEpChromeData): string | Promise<string> => {
        if (value.showTitle && (typeof value.title !== "string" || value.title.trim().length < 1)) {
            return "title cannot be empty";
        }
        if (value.showIcon && (typeof value.iconPath !== "string" || value.iconPath.trim().length < 1)) {
            return "icon cannot be empty";
        }
        if (this.props.onGetErrorMessage) {
            return this.props.onGetErrorMessage(value);
        } else {
            return '';
        }
    }
}
