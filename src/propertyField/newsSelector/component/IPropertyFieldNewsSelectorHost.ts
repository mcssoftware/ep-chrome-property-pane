import { IPropertyFieldNewsSelectorPropsInternal, IPropertyFieldNewsSelectorData } from "../IPropertyFieldNewsSelector";
import { ITermStore, ITerm, IGroup, ISPTermStorePickerService, ITermSet } from "../../../services/ISPTermStorePickerService";
import { IPickerTerms } from "../termStoreEntity";
import { IDropdownOption } from "office-ui-fabric-react/lib-es2015/Dropdown";

/**
 * PropertyFieldNewsSelectorHost properties interface
 */
export interface IPropertyFieldNewsSelectorHostProps extends IPropertyFieldNewsSelectorPropsInternal {
  onChange: (targetProperty?: string, newValue?: any) => void;
}

/**
 * PropertyFieldNewsSelectorHost state interface
 */
export interface IPropertyFieldNewsSelectorHostState {
  termStores?: ITermStore[];
  errorMessage?: string;
  openPanel?: boolean;
  termStoreLoaded?: boolean;
  pagesLoaded: boolean;
  activeValues: IPropertyFieldNewsSelectorData;
  pageDropDownOptions: IDropdownOption[];
  // activeNodes?: IPickerTerms;
}

export interface ITermChanges {
  changedCallback: (term: ITerm, termGroup: string, checked: boolean) => void;
  activeNodes?: IPickerTerms;
}

export interface ITermGroupProps extends ITermChanges {
  group: IGroup;
  termstore: string;
  termsService: ISPTermStorePickerService;
  multiSelection: boolean;
  isTermSetSelectable?: boolean;
  disabledTermIds?: string[];
}

export interface ITermGroupState {
  expanded: boolean;
  loaded?: boolean;
}

export interface ITermSetProps extends ITermChanges {
  termset: ITermSet;
  termstore: string;
  termGroup: string;
  termsService: ISPTermStorePickerService;
  autoExpand: () => void;
  multiSelection: boolean;
  isTermSetSelectable?: boolean;
  disabledTermIds?: string[];
}

export interface ITermSetState {
  terms?: ITerm[];
  loaded?: boolean;
  expanded?: boolean;
}

export interface ITermProps extends ITermChanges {
  termset: string;
  termGroup: string;
  term: ITerm;
  multiSelection: boolean;
  disabled: boolean;
}

export interface ITermState {
  selected?: boolean;
}
