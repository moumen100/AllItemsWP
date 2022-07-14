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
import {ISPListsItems, ISPLists} from './Helpers/SPHelpers'
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IBbsCpaAllItemsWebPartProps {
  description: string;
  seeAllUrl: string;
  listId: string;
  Items: IAllItems[];
  WebUrl: string;
  view: string;}

export default class BbsCpaAllItemsWebPart extends BaseClientSideWebPart<IBbsCpaAllItemsWebPartProps> {
  private lists: IPropertyPaneDropdownOption[] = [];
  private islistdropdowndisabled: boolean = false;
  private Helper: SPHelper;

  
  
  public render(): void {
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

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected  onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context,
    });
    this.Helper = new SPHelper(this.context,this.properties.WebUrl)
    return Promise.resolve();
  }
  public async _getListData(): Promise<ISPLists> {
    let url =
      !this.properties.WebUrl || this.properties.WebUrl == ""
        ? this.context.pageContext.web.absoluteUrl
        : this.properties.WebUrl;
    return this.context.spHttpClient
      .get(
        url + `/_api/web/lists?$filter=BaseTemplate eq 100`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
  public async getListItems(listId:string) : Promise< ISPListsItems> {

    let url =  (!this.properties.WebUrl || this.properties.WebUrl=='' ) ? this.context.pageContext.web.absoluteUrl  : this.properties.WebUrl
  url = url+ "/_api/Web/Lists(@v0)/Items?&@v0=guid'" + this.properties.listId +"'"
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }
  // public async GetSites(url: string): Promise<string> {
  //   const siteurl: string = await sp.site.getWebUrlFromPageUrl(url);
  //   return siteurl;
  // }
  protected async onPropertyPaneConfigurationStart() {
    this.lists = [];
    let listOfLists = await (await this._getListData()).value;
    listOfLists.forEach((elem) => {
      this.lists.push({ key: elem.Id, text: elem.Title });
    });
    this.lists.length > 1
      ? (this.islistdropdowndisabled = false)
      : (this.islistdropdowndisabled = true);
    this.properties.Items = this.properties.listId
      ? (
          await this.getListItems(
            this.properties.listId
          )
        ).value
      : new Array<IAllItems>();

    this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.render();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath == "WebUrl") {
      this.lists = [];
      let listOfLists = (await this._getListData()).value;
      listOfLists.forEach((elem) => {
        this.lists.push({ key: elem.Id, text: elem.Title });
      });
      this.lists.length > 1
        ? (this.islistdropdowndisabled = false)
        : (this.islistdropdowndisabled = true);
      this.properties.Items = this.properties.listId
        ? (
            await this.getListItems(
              this.properties.listId
            )
          ).value
        : new Array<IAllItems>();
      this.context.propertyPane.refresh();
      this.context.statusRenderer.clearLoadingIndicator(this.domElement);
      this.render();
    }
    if (propertyPath == "listId") {
      this.properties.Items = this.properties.listId
        ? (
            await this.getListItems(
              this.properties.listId
            )
          ).value
        : new Array<IAllItems>();
        this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.render();
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
                  value:this.context.pageContext.web.absoluteUrl
                }),

                PropertyPaneDropdown("listId", {
                  label: "Selectionner une liste",
                  options: [...this.lists],
                  selectedKey: this.properties.listId,
                  disabled: this.islistdropdowndisabled,
                
                }),

                PropertyPaneTextField("description", {
                  label: "Titre",
                  value:"Tout les produits" 
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
