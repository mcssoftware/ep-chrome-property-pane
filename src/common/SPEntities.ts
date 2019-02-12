/**
 * Represents SP List
 */
export interface ISPList {
    Id: string;
    Title: string;
    BaseTemplate: string;
}

/**
 * Replica of the returned value from the REST api
 */
export interface ISPLists {
    value: ISPList[];
}

