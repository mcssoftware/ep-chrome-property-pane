import { ISPService, IListPickerProps, ISPLists, ListPickerOrderByType } from "./ISPService";
import { SPHttpClient, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IWebPartContext } from "@microsoft/sp-webpart-base";

export default class SPService implements ISPService {
    constructor(private _context: IWebPartContext) { }

    public async getLists(properties: IListPickerProps, webUrl?: string): Promise<ISPLists> {
        try {
            const webAbsoluteUrl = !webUrl ? this._context.pageContext.web.absoluteUrl : webUrl;
            let queryUrl: string = `${webAbsoluteUrl}/_api/lists?$select=Title,id,BaseTemplate`;
            // Check if the orderBy property is provided
            if (typeof properties.listOrderBy !== "undefined" && properties.listOrderBy !== null) {
                queryUrl += '&$orderby=';
                switch (properties.listOrderBy) {
                    case ListPickerOrderByType.Id:
                        queryUrl += 'Id';
                        break;
                    case ListPickerOrderByType.Title:
                        queryUrl += 'Title';
                        break;
                }
            }
            // Check if the list have get filtered based on the list base template type
            if (properties.listBaseTemplate !== null && properties.listBaseTemplate) {
                queryUrl += '&$filter=BaseTemplate%20eq%20';
                queryUrl += properties.listBaseTemplate;
                // Check if you also want to exclude hidden list in the list
                if (properties.includeHiddenList === false) {
                    queryUrl += '%20and%20Hidden%20eq%20false';
                }
            } else {
                if (properties.includeHiddenList === false) {
                    queryUrl += '&$filter=Hidden%20eq%20false';
                }
            }
            const data = await this._context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1);
            if (data.ok) {
                const results = await data.json();
                return results as ISPLists;
            }
            return {
                value: []
            };
        } catch (error) {
            return Promise.reject(error);
        }
    }

    /**
   * Get List Items
   */
    public async getListItems(filterText: string, listTitle: string, internalColumnName: string, webUrl?: string, orderBy?: string, top?: number): Promise<any[]> {
        try {
            const webAbsoluteUrl = !webUrl ? this._context.pageContext.web.absoluteUrl : webUrl;
            const numberOfItems: number = top || 100;
            const apiUrl = `${webAbsoluteUrl}/_api/web/lists/GetByTitle('${listTitle}')/items?$select=Id,${internalColumnName}` +
                `&$filter=${filterText}&$order=${orderBy || ""}&$top=${numberOfItems}`;
            const data = await this._context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
            if (data.ok) {
                const results = await data.json();
                if (results && results.value && results.value.length > 0) {
                    return results.value;
                }
            }
            return [];
        } catch (error) {
            return Promise.reject(error);
        }
    }

    /**
  * Get List Items using caml query
  */
    public async getListItemsCaml?(camlQuery: string, listTitle: string, webUrl?: string): Promise<any[]> {
        try {
            const webAbsoluteUrl = !webUrl ? this._context.pageContext.web.absoluteUrl : webUrl;
            const apiUrl = `${webAbsoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/GetItems(query=@v1)?@v1={\"ViewXml\":\"${camlQuery}\"}`;
            const spOpts: ISPHttpClientOptions = {
                body: ""
            };
            const data = await this._context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts);
            if (data.ok) {
                const results = await data.json();
                if (results && results.value && results.value.length > 0) {
                    return results.value;
                }
            }
            return [];
        } catch (error) {
            return Promise.reject(error);
        }
    }
}