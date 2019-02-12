/**
 * Selected terms
 */
export interface IPickerTerm {
    name: string;
    key: string;
    path: string;
    termSet: string;
    termSetName?: string;
    labels?: string[];
    termGroup: string;
}

export interface IPickerTerms extends Array<IPickerTerm> { }

/**
 * Generic Term Object (abstract interface)
 */
export interface ISPTermObject {
    Name: string;
    Guid: string;
    Identity: string;
    leaf: boolean;
    children?: ISPTermObject[];
    collapsed?: boolean;
    type: string;
}

/**
 * Defines a SharePoint Term Store
 */
export interface ISPTermStore extends ISPTermObject {
    IsOnline: boolean;
    WorkingLanguage: string;
    DefaultLanguage: string;
    Languages: string[];
}

/**
 * Defines an array of Term Stores
 */
export interface ISPTermStores extends Array<ISPTermStore> {
}

/**
 * Defines a Term Store Group of term sets
 */
export interface ISPTermGroup extends ISPTermObject {
    IsSiteCollectionGroup: boolean;
    IsSystemGroup: boolean;
    CreatedDate: string;
    LastModifiedDate: string;
}

/**
 * Array of Term Groups
 */
export interface ISPTermGroups extends Array<ISPTermGroup> {
}
