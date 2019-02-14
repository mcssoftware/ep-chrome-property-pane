/**
 * Enum for specifying how the lists should be sorted
 */
export enum ListPickerOrderByType {
    Id = 1,
    Title
}

export interface IListPickerProps {
    /**
    * BaseTemplate ID of the lists or libaries you want to return.
    */
    listBaseTemplate?: number;
    /**
     * Specify if you want to include or exclude hidden lists. By default this is true.
     */
    includeHiddenList?: boolean;
    /**
     * Specify the property on which you want to order the retrieve set of lists.
     */
    listOrderBy?: ListPickerOrderByType;
}

/**
 * Defines a SharePoint list
 */
export interface ISPList {
    Title: string;
    Id: string;
    BaseTemplate: string;
}

/**
 * Defines a collection of SharePoint lists
 */
export interface ISPLists {
    value: ISPList[];
}


export interface ISPService {
    getLists(properties: IListPickerProps, webUrl?: string): Promise<ISPLists>;
    getListItems?(filterText: string, listTitle: string, internalColumnName: string, webUrl?: string, orderBy?: string, top?: number): Promise<any[]>;
    getListItemsCaml?(camlQuery: string, listTitle: string, webUrl?: string): Promise<any[]>;
}