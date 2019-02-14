import { IPropertyFieldCalendarData, IPropertyFieldCalendarPropsInternal } from "../IPropertyFieldCalendar";
import { IDropdownOption } from "office-ui-fabric-react/lib-es2015/Dropdown";

export interface IPropertyPaneCalendarHostProps extends IPropertyFieldCalendarPropsInternal {
    targetProperty: string;
    /**
    * Callback for the onChanged event.
    */
    onChange: (targetProperty?: string, newValue?: IPropertyFieldCalendarData) => void;
}

export interface IPropertyPaneCalendarHostState {
    value: IPropertyFieldCalendarData;
    errorMessage?: string;
    listOptions: IDropdownOption[];
    listLoaded: boolean;
    itemsDropDownOptions: IDropdownOption[];
    itemsLoaded: boolean;
}