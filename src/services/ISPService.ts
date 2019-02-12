export interface ISPService {
    getListItems?(filterText: string, listTitle: string, internalColumnName: string, webUrl?: string): Promise<any[]>;
}