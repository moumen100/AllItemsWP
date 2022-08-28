import { IAllItems,  IAllItemsList } from "../Models/IAllItems";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {sp} from "@pnp/sp"

/**
 * Interface pour définir le résulatat de API Liste des listes 
 */
export interface ISPLists {
  value: IAllItemsList[];
}
/**
 * une interface pour définir le résultat de API liste des Items 
 */
export interface ISPListsItems {
  value: IAllItems[];
}
/**
 * Cette class contient des méthodes qui consumme Rest API sharepoint afin de récupéré les données de puis la colection de site 
 */
export default class SPHelper {
  private context: WebPartContext; // attribut  contient le context sharepoint récupérer depuis spfx framwork
  private WebUrl: string;  //attribut pour sauvegarder la valeur WebUrl depuis le configuration de proprety pane 
  /**
   * constructeur de l'objet 
   * @param cnt  l'objet context nous pemettre d'utiliser les ressources sharepoint  
   * @param WebUrl l'url de la collection de site configuré sur la proprety pane 
   */
  public constructor(cnt: WebPartContext, WebUrl) { 
    this.context = cnt;
    this.WebUrl = WebUrl;
  }
  /**
   * Méthode permet de consumer l'API sharepoint pour récupérer les listes sharepoint depuis la collection de site configuré 
   * @returns  Objet de type ISPLists qui contient comme prorety array de type IAllItemsList qui contient liste des produits 
   */
  public _getListData(): Promise<ISPLists> {
    // récupérer la url de host de la web application approprié depuis la configuration de la proprety pane 
    let url =  (!this.WebUrl || this.WebUrl=='' ) ? this.context.pageContext.web.absoluteUrl  : this.WebUrl
    // construire l'url   de end point de l'API sharepoint  utilisé 
    // On utilise la condition BaseTemplate eq 100 afin de récupérer uniquement les custom list
     url +=
      `/_api/web/lists?$filter=BaseTemplate eq 100`;
      //consumer l'api indiqué sur l'objet url 
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  /**
   * Méthode Permet de récupéré la liste des Items(Produits) récupéré depuis l'PI Sharepoint 
   * @param listId  ID de la liste sharepoint selectioné sur la proprety pane 
   * @returns  retourner la liste des Items (produits) de type ISPListsItems qui contient une proprety array de type IAllItems

   */
  public async getListItems(listId:string) : Promise< ISPListsItems> {
        // récupérer la url de host de la web application approprié depuis la configuration de la proprety pane 
    let url =  (!this.WebUrl || this.WebUrl=='' ) ? this.context.pageContext.web.absoluteUrl  : await this.WebUrl
    //construire url de APi sharepoint à utilisé  avec paramètre Id de la liste 
    // on ajoute le paramètre active ne true Afin de récupérer uniquement les items activé 
  url += `/_api/Web/Lists(@v0)/Items?&@v0=guid'${listId}'&$filter=active ne true`
  //consumer l'API pour récupéré les valeurs 
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
}
