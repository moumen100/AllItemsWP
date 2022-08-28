import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'BbsCpaAllItemsWebPartStrings';
import BbsCpaAllItems from './components/BbsCpaAllItems';
import { IBbsCpaAllItemsProps } from './components/IBbsCpaAllItemsProps';
import { IAllItems } from './Models/IAllItems';
import SPHelper from './Helpers/SPHelpers';
import { sp } from '@pnp/sp'; 

export interface IBbsCpaAllItemsWebPartProps {
  description: string;
  seeAllUrl: string;
  listId: string;
  Items: IAllItems[];
  WebUrl: string;
  view: string;}

export default class BbsCpaAllItemsWebPart extends BaseClientSideWebPart<IBbsCpaAllItemsWebPartProps> {
  private lists: IPropertyPaneDropdownOption[] = []; // tableau pour sauvegarder la listes des listes dans le champs select List 
  private islistdropdowndisabled: boolean = false; // condition pour activer/déactiver le champs selectList , true si la liste des liste est vide 

  
  /**
   * Méthode pour render la webpart en initialisant le copone nt react 
   */
  public render(): void {
    //initiaiser les valeurs des props react 
    const element: React.ReactElement<IBbsCpaAllItemsProps > = React.createElement(
      BbsCpaAllItems,
      {
        description: this.properties.description,
        listitem: this.properties.Items,
        Context: this.context,
        SeeAllUrl: this.properties.seeAllUrl,
        listId: this.properties.listId,
        WebUrl: this.properties.WebUrl,
        view: this.properties.view,
      }
    );
    //rende  le component react 
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected  onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context,
    });
    return Promise.resolve();
  }
 
  /**
   * évènement utilisé par spfx se déclanche lors de lancement la proprety pane 
   */
  protected async onPropertyPaneConfigurationStart() {
    this.lists = []; // initialiser la liste des listes 
    let helpers : SPHelper = new SPHelper(this.context,this.properties.WebUrl) // créer l'objet helper afin d'untiliser les méthoes de la class SpHelper
    let listOfLists = await (await helpers._getListData()).value; // récupérer la liste des listes sharepoint avec la méthodes _getListData qui utilise API Sharepoint  
    listOfLists.forEach((elem) => {
      this.lists.push({ key: elem.Id, text: elem.Title }); // mettre en places la liste des liste dans le champs dropdown Select List 
    });
    // activer/déactiver le champs suite au résultat obtenu Si la liste est vide le champs  dropdownsera déactiver 
    this.lists.length > 1
      ? (this.islistdropdowndisabled = false)
      : (this.islistdropdowndisabled = true);
    // refresh la propretypane 
    this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.render(); // refrsh le coponent react 
  }
/**
 * évènement par spfx se déclanche lors de modification d'un champs dans la propretypane 
 * @param propertyPath  paramètre qui indique le champs modifié 
 * @param oldValue l'ancienne valeur de champs 
 * @param newValue la nouvelle valeur de champs 
 */
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    let helpers : SPHelper = new SPHelper(this.context,this.properties.WebUrl) // créer l'objet helper afin d'untiliser les méthoes de la class SpHelper
    // si le champs WebUrl change on change les valeurs de champs selectList la liste des listes sharepoint 
    if (propertyPath == "WebUrl") { 
      this.lists = []; // initiliser le liste des listes 
      let listOfLists = (await helpers._getListData()).value;  // récupérer la nouvelle  liste des listes sharepoint avec la méthodes _getListData qui utilise API Sharepoint  
      listOfLists.forEach((elem) => {
        this.lists.push({ key: elem.Id, text: elem.Title }); // mettre en place le nouvelle liste des listes sharepoint dans le champs dropdown 
      });
          // activer/déactiver le champs suite au résultat obtenu Si la liste est vide le champs  dropdownsera déactiver 

      this.lists.length > 1
        ? (this.islistdropdowndisabled = false)
        : (this.islistdropdowndisabled = true);
        //refresh la proprety pane 
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render(); // refresh componet react 
    }
  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("WebUrl", {
                  label: "Collection de site",
                }), // champs WebUrl configure la collection de site source des donnés 

                PropertyPaneDropdown("listId", {
                  label: "Selectionner une liste",
                  options: [...this.lists],
                  selectedKey: this.properties.listId,
                  disabled: this.islistdropdowndisabled,
                }), // Champs Selectionner une liste sharepoint depuis la liste des lists 

                PropertyPaneTextField("description", {
                  label: "Titre", 
                  value:"Tout les produits" 
                }) // champs Titre configure le titre de la webpart 
              ]
            }
          ]
        }
      ]
    };
  }
}
