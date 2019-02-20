import * as React from "react";
import { IContentData, getContentDataDefaultValue } from "../IPropertyPaneMultiZoneSelector";
import styles from "./PropertyFieldMultiZoneSelectorHost.module.scss";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { cloneDeep } from "@microsoft/sp-lodash-subset";
import { ColorPicker } from "office-ui-fabric-react/lib/ColorPicker";
import { Label } from "office-ui-fabric-react/lib/Label";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";

export interface IContentControlProps {
    data: IContentData;
}

export interface IContentControlState {
    value: IContentData;
    useImage: boolean;
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
        this.state = {
            value,
            useImage: false
        };
    }

    /**
     * Render content control
     * @returns {JSX.Element}
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
     */
    public render(): JSX.Element {
        const { value, useImage } = this.state;
        return (
            <div className={styles.propertyFieldMultiZoneNewsSelectorHost}>
                <div className={styles.row}>
                    <div className={styles.column}>
                        <Toggle label="Background Color"
                            onText="Use Image as background"
                            offText="Use color as background"
                            onChanged={this.onBackgroundToggled}
                            checked={useImage} />
                    </div>
                    <div className={styles.column + " " + styles.bgwrapper}>
                        {!useImage && <div>
                            <ColorPicker color={value.backgroundColor}
                                onColorChanged={this.onColorChanged}
                                alphaSliderHidden={true} />
                            <div className={styles.column2}>
                                <div
                                    className={styles.bgcolorSquare}
                                    style={{
                                        backgroundColor: value.backgroundColor
                                    }}
                                />
                            </div>
                        </div>}
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
                            onChanged={this.ontargetUrlTextChanged} />
                    </div>
                </div>
            </div>
        );
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
     * Color selected.
     * @private
     * @memberof ContentControl
     */
    private onColorChanged = (color: string): void => {
        const value = cloneDeep(this.state.value);
        value.backgroundColor = color;
        this.setState({ value });
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
    }
}