import { ISPService } from "./ISPService";
import { SPHttpClient } from "@microsoft/sp-http";
import { IWebPartContext } from "@microsoft/sp-webpart-base";

export default class SPService implements ISPService {
    constructor(private _context: IWebPartContext ) { }

    /**
   * Get List Items
   */
    public async getListItems(filterText: string, listTitle: string, internalColumnName: string, webUrl?: string): Promise<any[]> {
        try {
            const webAbsoluteUrl = !webUrl ? this._context.pageContext.web.absoluteUrl : webUrl;
            const apiUrl = `${webAbsoluteUrl}/_api/web/lists/GetByTitle('${listTitle}')/items?$select=Id,${internalColumnName}&$filter='${filterText}'`;
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
}