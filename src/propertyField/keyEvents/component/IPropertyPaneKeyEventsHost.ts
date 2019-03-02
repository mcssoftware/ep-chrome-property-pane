import { IPropertyFieldKeyEventsPropsInternal, IPropertyPaneKeyEventsData } from "../IPropertyPaneKeyEvents";
import { IDropdownOption } from "office-ui-fabric-react/lib-es2015/Dropdown";

export interface IPropertyPaneKeyEventsHostProps extends IPropertyFieldKeyEventsPropsInternal {
    targetProperty: string;
    /**
    * Callback for the onChanged event.
    */
    onChange: (targetProperty?: string, newValue?: IPropertyPaneKeyEventsData) => void;
}

export interface IPropertyPaneKeyEventsHostState {
    value: IPropertyPaneKeyEventsData;
    listOptions: IDropdownOption[];
    listLoaded: boolean;
    listChoiceKey: string;
    selectedLists: string[];
    errorMessage?: string;
    numberOfItemsText: string;
}