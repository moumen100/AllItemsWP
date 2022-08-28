import { IAllItems,  IAllItemsList } from "../Models/IAllItems";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {sp} from "@pnp/sp"
export interface ISPLists {
  value: IAllItemsList[];
}

export interface ISPListsItems {
  value: IAllItems[];
}

export default class SPHelper {
  private context: WebPartContext;
  private WebUrl: string;
  public constructor(cnt: WebPartContext, WebUrl) {
    this.context = cnt;
    this.WebUrl = WebUrl;
  }
  // liste des listes de sharepoint
  public _getListData(): Promise<ISPLists> {
    let url =  (!this.WebUrl || this.WebUrl=='' ) ? this.context.pageContext.web.absoluteUrl  : this.WebUrl

     url +=
      `/_api/web/lists?$filter=BaseTemplate eq 100`;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  public async GetSites(url: string): Promise<string> {
    const siteurl: string = await sp.site.getWebUrlFromPageUrl(url);
    return siteurl;
  }

  public async getListItems(listId:string) : Promise< ISPListsItems> {

    let url =  (!this.WebUrl || this.WebUrl=='' ) ? this.context.pageContext.web.absoluteUrl  : await this.WebUrl
  url += "/_api/Web/Lists(@v0)/Items?&@v0=guid'" + listId +"'"
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
}
