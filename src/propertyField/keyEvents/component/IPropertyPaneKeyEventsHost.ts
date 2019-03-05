import { IPropertyFieldKeyEventsPropsInternal, IPropertyFieldKeyEventsData } from "../IPropertyFieldKeyEvents";
import { IDropdownOption } from "office-ui-fabric-react/lib-es2015/Dropdown";

export interface IPropertyPaneKeyEventsHostProps extends IPropertyFieldKeyEventsPropsInternal {
    targetProperty: string;
    /**
    * Callback for the onChanged event.
    */
    onChange: (targetProperty?: string, newValue?: IPropertyFieldKeyEventsData) => void;
}

export interface IPropertyPaneKeyEventsHostState {
    value: IPropertyFieldKeyEventsData;
    listOptions: IDropdownOption[];
    listLoaded: boolean;
    listChoiceKey: string;
    selectedLists: string[];
    errorMessage?: string;
    numberOfItemsText: string;
}