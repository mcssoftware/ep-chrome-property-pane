export interface ISPService {
    getListItems?(filterText: string, listTitle: string, internalColumnName: string, webUrl?: string): Promise<any[]>;
    getListItemsCaml?(camlQuery: string, listTitle: string, webUrl?: string): Promise<any[]>;
}