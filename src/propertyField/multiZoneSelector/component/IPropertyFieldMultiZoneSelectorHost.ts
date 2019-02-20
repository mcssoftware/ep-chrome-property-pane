import * as React from "react";
import {
    IPropertyFieldMultiZoneSelectorPropsInternal as IPropertyFieldMultiZoneSelectorPropsInternal,
    IPropertyPaneMultiZoneSelectorData,
    IZoneData} from "../IPropertyPaneMultiZoneSelector";
import { ZoneDataHost } from "./ZoneDataHost";

/**
 * PropertyFieldNewsSelectorHost properties interface
 */
export interface IPropertyFieldMultiZoneSelectorHostProps extends IPropertyFieldMultiZoneSelectorPropsInternal {
    targetProperty: string;
    /**
    * Callback for the onChanged event.
    */
    onChange: (targetProperty?: string, newValue?: IPropertyPaneMultiZoneSelectorData) => void;
}

export interface IPropertyFieldMultiZoneSelectorHostState {
    zoneSelected: number;
    errorMessage?: string;
    activeValues: ZoneDataHost[];
    selectedZoneData: ZoneDataHost;
}