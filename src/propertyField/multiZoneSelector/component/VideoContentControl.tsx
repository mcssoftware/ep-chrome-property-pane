import * as React from "react";
import { IVideoData, getVideoDataDefaultValue } from "../IPropertyPaneMultiZoneSelector";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import styles from "./PropertyFieldMultiZoneSelectorHost.module.scss";

export interface IVideoContentControlProps {
    data: IVideoData;
    /**
     * Defines a onPropertyChange function to raise when the selected value changed.
     * Normally this function must be always defined with the 'this.onPropertyChange'
     * method of the web part object.
     */
    notify(oldValue: IVideoData, newValue: IVideoData): void;
    /**
     * Whether the property pane field is enabled or not.
     */
    disabled?: boolean;
}

export interface IVideoContentControlState {
    value: IVideoData;
}

/**
 * Control used by video content type
 * @export
 * @class VideoContentControl
 * @extends {React.Component<IVideoContentControlProps, IVideoContentControlState>}
 */
export class VideoContentControl extends React.Component<IVideoContentControlProps, IVideoContentControlState>{
    constructor(props: IVideoContentControlProps) {
        super(props);
        this.state = {
            value: typeof props.data !== "undefined" ? cloneDeep(props.data) : getVideoDataDefaultValue(),
        };
    }

    public render(): JSX.Element {
        const { value } = this.state;
        const forcedDisabled: boolean = this.props.disabled || false;
        return (
            <div className={styles.propertyFieldMultiZoneNewsSelectorHost}>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <TextField label="Video Url"
                            value={value.url}
                            disabled={forcedDisabled}
                            onChanged={this.onVideoUrlChanged} />
                    </div>
                </div>
            </div>
        );
    }

    private onVideoUrlChanged = (newValue: string): void => {
        const value = cloneDeep(this.state.value);
        value.url = newValue;
        this.setState({ value });
        this.validate(value);
    }

    /**
     * Validates the new custom field value
     * @private
     * @param {IPropertyPaneMultiZoneSelectorData} value
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
    */
    private validate = (newValue: IVideoData): void => {
        const internalResult: string = this.validateInternal(newValue);
        if (internalResult.length === 0) {
            this.notifyAfterValidate(newValue);
        }
    }

    private validateInternal = (newValue: IVideoData): string => {
        return "";
    }

    private notifyAfterValidate = (newValue: IVideoData) => {
        if (typeof this.props.notify === "function") {
            this.props.notify(this.props.data, newValue);
        }
    }
}
