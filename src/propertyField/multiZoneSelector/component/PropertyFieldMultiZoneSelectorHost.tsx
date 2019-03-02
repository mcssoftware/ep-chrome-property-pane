import * as React from "react";
import { Async } from "office-ui-fabric-react/lib/Utilities";
import { IPropertyFieldMultiZoneSelectorHostProps, IPropertyFieldMultiZoneSelectorHostState } from "./IPropertyFieldMultiZoneSelectorHost";
import FieldErrorMessage from "../../errorMessage/FieldErrorMessage";
import {
    IPropertyPaneMultiZoneSelectorData,
    ZoneDataType,
    IContentData,
    IVideoData
} from "../IPropertyPaneMultiZoneSelector";
import { initGlobalVars } from "../../../common/ep";
import { IChoiceGroupOption, ChoiceGroup } from "office-ui-fabric-react/lib/ChoiceGroup";
import styles from "./PropertyFieldMultiZoneSelectorHost.module.scss";
import Header from "../../header/header";
import PropertyFieldNewsSelectorHost from "../../newsSelector/component/PropertyFieldNewsSelectorHost";
import { Label } from "office-ui-fabric-react/lib/Label";
import { ContentControl } from "./ContentControl";
import { VideoContentControl } from "./VideoContentControl";
import { ZoneDataHost } from "./ZoneDataHost";
import { cloneDeep } from "@microsoft/sp-lodash-subset";

const images: string[] = [
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAIAAAAt/+nTAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACbSURBVFhH7dgrDgQhEEXRaRDl2AmS4BA42A2W5eEIK0LNJ6mMB8QLSR1VPclL+poW8xhjvPdzztc6ImqthRBQ8977k1KqtfJv60opYwx+WOecO5nnnNVe/R98rvi8lgSgSQCaBKBJAJoEoCmtNZ9brp4L8RNjfB+w1vK15XD+fXn5jKJJAJoEoEkAmgSg3R9ARHxugc8v/3u99w9GNPBvnwpnmAAAAABJRU5ErkJggg==",
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAIAAAAt/+nTAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACgSURBVFhH7ditDQQhEIbhWxDj6ARJcAgcdEMl2w0F4AgVoe4nmZyHFV9I5lEDySS8ZsVexhjv/ZzztY6IWmshBNR67/1KKdVa+W6dc+6+bz6sK6WMMfiwLues9ur/4OuKx2NJAJoEoEkAmgSgSQCa0lrzuOXodSF+YozvB6y1PG15uP59vHxG0SQATQLQJABNAtDODyAiHrfA1w//vd77B0Y08G+lGp7jAAAAAElFTkSuQmCC",
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAIAAAAt/+nTAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACeSURBVFhH7dghDsQgEIXhLYhx3ARJcAgcF8NyGw6AI5wI1e0mk/VQ8UIynxpIJuE3Fb2MMd77OednHRG11kIIqPXe+5VSqrXy3Trn3BiDD+ue9VIKH9blnNVe/R98XfF4LAlAkwA0CUCTADQJQFNaax63HL0uxE+M8X7BWsvTlpfrz+PlM4omAWgSgCYBaBKAdn4AEfG4Bb5++O/13r9GNPBvkUR9AgAAAABJRU5ErkJggg==",
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAIAAAAt/+nTAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACcSURBVFhH7dghDsUgEEXRFsQ4loIiwSFAsRV2wdZwWHaD6q+YfA8VLyRz1FS8pNdU9DbGeO/nnNc6ImqthRBQ8977lXN+PnDO8bXl4/x9ebVX/wefKz6PJQFoEoAmAWgSgCYBaEprzeeWo+dCvO4YY62Vn9aVUsYY/LDOWvtlnlKSzyiaBKBJAJoEoEkA2vkBRMTnFvj88N/rvf8AcMTwrjiLhyYAAAAASUVORK5CYII=",
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAIAAAAt/+nTAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACfSURBVFhH7dghDsUgEEXRFgSOpaBIcAhQbIUFsRsMDstuUP0Vk++h4oVkjpoRk/SaCm6ttXNuznmtU0q11rz3qPPe+5VSej6w1tK05eP5+/Fir/4Pfi5oPBYHoHEAGgegcQAaB6AJKSWNW44+Z+x1hxBqrbStM8aUUmhZl3MeY9CyLsbIv1E0DkDjADQOQOMAtPMDlFI0boGfH/683vsPcMTwrkq5W9wAAAAASUVORK5CYII=",
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAEAAAAAgCAIAAAAt/+nTAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAACeSURBVFhH7dghDsYgDIbhDUQdR5kiwSFwnIpbcBsMDsttUPsRze9h4gtJH9WKJnvNBLcxxjk3xrjWEVGt1XuPOm+tXTHG9wNrLU9bPp7Pj1d79X/wc8XjsSQATQLQJABNAtAkAE1prXnccvS5ENMdQiil8LbueZ7eOy/r5nnOmZd1KSX5jaJJAJoEoEkAmgSgnR9ARDxugZ8f/rze2g9wxPCurcet7AAAAABJRU5ErkJggg==",
];

export class PropertyFieldMultiZoneNewsSelectorHost extends React.Component<IPropertyFieldMultiZoneSelectorHostProps, IPropertyFieldMultiZoneSelectorHostState> {
    private async: Async;
    private delayedValidate: (value: IPropertyPaneMultiZoneSelectorData) => void;
    private zoneOptions: IChoiceGroupOption[];

    /**
     * Constructor method
     */
    constructor(props: IPropertyFieldMultiZoneSelectorHostProps) {
        super(props);
        if (typeof (window as any).Epmodern === "undefined") {
            initGlobalVars();
        }
        this.zoneOptions = [];
        const activeValues: ZoneDataHost[] = [];

        for (let index: number = 0; index < this.props.numberOfZones; index++) {
            const text: string = "Zone " + (index + 1).toString();
            this.zoneOptions.push({
                key: index.toString(),
                imageSrc: images[index],
                imageAlt: text,
                selectedImageSrc: images[index],
                imageSize: { width: 32, height: 32 },
                text: text
            });
            activeValues.push(new ZoneDataHost(this.props.value[index]));
        }

        this.state = {
            activeValues,
            errorMessage: this.validateInternal(activeValues),
            zoneSelected: this.zoneOptions.length > 0 ? 0 : -1,
            selectedZoneData: activeValues[0]
        };
        this.async = new Async(this);
        this.delayedValidate = this.async.debounce(this.validate, this.props.deferredValidationTime);
    }

    /**
     * Render multizone news selector with Office UI  Fabric
     * @returns {JSX.Element}
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
     */
    public render(): JSX.Element {
        const { selectedZoneData, errorMessage } = this.state;
        const selectedZoneDataType = selectedZoneData.getType().toString();
        const forcedDisabled: boolean = this.props.disabled || false;
        return (
            <div className={styles.propertyFieldMultiZoneNewsSelectorHost}>
                <Header title={this.props.label} />
                <div className={styles.row}>
                    <div className={styles.column}>
                        <ChoiceGroup label="Select zone"
                            selectedKey={this.state.zoneSelected.toString()}
                            options={this.zoneOptions}
                            disabled={forcedDisabled}
                            onChange={this.onZoneOptionSelected} />
                    </div>
                </div>
                {(selectedZoneData !== null) &&
                    <div>
                        <div className={styles.row}>
                            <div className={styles.column}>
                                <ChoiceGroup label="Select zone data type"
                                    selectedKey={selectedZoneDataType}
                                    disabled={forcedDisabled}
                                    options={[
                                        {
                                            key: ZoneDataType.Content.toString(),
                                            iconProps: { iconName: "InsertTextBox" },
                                            text: "Content",
                                        },
                                        {
                                            key: ZoneDataType.Video.toString(),
                                            iconProps: { iconName: "Video" },
                                            text: "Video"
                                        },
                                        {
                                            key: ZoneDataType.Article.toString(),
                                            iconProps: { iconName: "Articles" },
                                            text: "Article",
                                        }
                                    ]}
                                    onChange={this.onArticleTypeSelected} />
                            </div>
                        </div>
                        <div className={styles.row}>
                            <div className={styles.column}>
                                {selectedZoneData.getType() === ZoneDataType.Article && <PropertyFieldNewsSelectorHost
                                    showHeader={false}
                                    allowMultipleSelections={this.props.allowMultipleSelections}
                                    excludeSystemGroup={this.props.excludeSystemGroup}
                                    limitByGroupNameOrID={this.props.limitByGroupNameOrID}
                                    limitByTermsetNameOrID={this.props.limitByTermsetNameOrID}
                                    hideTermStoreName={this.props.hideTermStoreName}
                                    isTermSetSelectable={this.props.isTermSetSelectable}
                                    context={this.props.context}
                                    onDispose={this.props.onDispose}
                                    onRender={this.props.onRender}
                                    disabled={this.props.disabled}
                                    onGetErrorMessage={null}
                                    deferredValidationTime={this.props.deferredValidationTime}
                                    termService={this.props.termService}
                                    spService={this.props.spService}
                                    value={selectedZoneData.getData() as any}
                                    panelTitle={this.props.panelTitle}
                                    targetProperty={this.props.targetProperty}
                                    label={this.props.label}
                                    onChange={null}
                                    onPropertyChange={this.onArticleDataChanged}
                                    key={this.props.key}
                                />}
                                {selectedZoneData.getType() === ZoneDataType.Content &&
                                    <ContentControl data={selectedZoneData.getData() as IContentData}
                                        disabled={forcedDisabled}
                                        notify={this.onContentDataChanged} />
                                }
                                {selectedZoneData.getType() === ZoneDataType.Video &&
                                    <VideoContentControl data={selectedZoneData.getData() as IVideoData}
                                        disabled={forcedDisabled}
                                        notify={this.onVideoDataChanged} />
                                }
                            </div>
                        </div>
                    </div>}
                {(selectedZoneData === null) && <div className={styles.row}>
                    <div className={styles.column}>
                        <Label>Invalid data</Label>
                    </div>
                </div>}
                <FieldErrorMessage errorMessage={errorMessage} />
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
     * Zone changed
     * @private
     * @param {React.SyntheticEvent<HTMLElement>} ev
     * @param {IChoiceGroupOption} option
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
     */
    private onZoneOptionSelected = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => {
        const zoneSelected = parseInt(option.key);
        const selectedZoneData = cloneDeep(this.state.activeValues[zoneSelected]);
        this.setState({
            zoneSelected,
            selectedZoneData
        });
    }

    /**
     * Article type changed
     * @private
     * @param {React.SyntheticEvent<HTMLElement>} ev
     * @param {IChoiceGroupOption} option
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
     */
    private onArticleTypeSelected = (ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption): void => {
        const selectedZoneData = cloneDeep(this.state.selectedZoneData);
        selectedZoneData.setZoneType(option.key);
        const activeValues = cloneDeep(this.state.activeValues);
        activeValues[this.state.zoneSelected] = selectedZoneData;
        this.setState({
            selectedZoneData,
            activeValues,
        });
    }

    /**
     * Callback from content control value on change
     * @private
     * @param {IContentData} oldValue
     * @param {IContentData} newValue
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
     */
    private onContentDataChanged = (oldValue: IContentData, newValue: IContentData): void => {
        const { zoneSelected, activeValues } = this.state;
        const selectedValue = cloneDeep(this.state.selectedZoneData);
        selectedValue.setValues({ type: ZoneDataType.Content, data: newValue });
        activeValues[zoneSelected] = selectedValue;
        this.setState({ activeValues, selectedZoneData: selectedValue });
        this.validate(activeValues);
    }

    /**
     * Callback from video control value on change
     *
     * @private
     * @param {IVideoData} oldValue
     * @param {IVideoData} newValue
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
     */
    private onVideoDataChanged = (oldValue: IVideoData, newValue: IVideoData): void => {
        const { zoneSelected, activeValues } = this.state;
        const selectedValue = cloneDeep(this.state.selectedZoneData);
        selectedValue.setValues({ type: ZoneDataType.Video, data: newValue });
        activeValues[zoneSelected] = selectedValue;
        this.setState({ activeValues, selectedZoneData: selectedValue });
        this.validate(activeValues);
    }

    /**
     * Callback from article control value on change
     *
     * @private
     * @param {string} targetProperty
     * @param {IVideoData} oldValue
     * @param {IVideoData} newValue
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
     */
    private onArticleDataChanged = (targetProperty: string, oldValue: IVideoData, newValue: IVideoData): void => {
        const { zoneSelected, activeValues } = this.state;
        const selectedValue = cloneDeep(this.state.selectedZoneData);
        selectedValue.setValues({ type: ZoneDataType.Article, data: newValue });
        activeValues[zoneSelected] = selectedValue;
        this.setState({ activeValues, selectedZoneData: selectedValue });
        this.validate(activeValues);
    }

    /**
     * Validates the new custom field value
     * @private
     * @param {IPropertyPaneMultiZoneSelectorData} value
     * @memberof PropertyFieldMultiZoneNewsSelectorHost
    */
    private validate = (value: ZoneDataHost[]): void => {
        const internalResult: string = this.validateInternal(value);
        if (internalResult.length === 0) {
            this.notifyAfterValidate(value);
        }
    }

    private validateInternal = (value: ZoneDataHost[]): string => {
        return "";
    }

    private notifyAfterValidate = (newvalues: ZoneDataHost[]) => {
        const newValue = newvalues.map(a => a.getZoneData());
        if (this.props.onPropertyChange && newValue.length > 0) {
            this.props.onPropertyChange(this.props.targetProperty, this.props.value, newValue);
            // Trigger the apply button
            if (typeof this.props.onChange !== "undefined" && this.props.onChange !== null) {
                this.props.onChange(this.props.targetProperty, newValue);
            }
        }
    }
}