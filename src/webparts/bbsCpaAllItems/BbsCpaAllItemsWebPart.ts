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
  private lists: IPropertyPaneDropdownOption[] = [];
  private islistdropdowndisabled: boolean = false;

  
  
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
    return Promise.resolve();
  }
 
 
  protected async onPropertyPaneConfigurationStart() {
    this.lists = [];
    let helpers : SPHelper = new SPHelper(this.context,this.properties.WebUrl)
    let listOfLists = await (await helpers._getListData()).value;
    listOfLists.forEach((elem) => {
      this.lists.push({ key: elem.Id, text: elem.Title });
    });
    this.lists.length > 1
      ? (this.islistdropdowndisabled = false)
      : (this.islistdropdowndisabled = true);

    this.context.propertyPane.refresh();
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.render();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    let helpers : SPHelper = new SPHelper(this.context,this.properties.WebUrl)
    if (propertyPath == "WebUrl") {
      this.lists = [];
      let listOfLists = (await helpers._getListData()).value;
      listOfLists.forEach((elem) => {
        this.lists.push({ key: elem.Id, text: elem.Title });
      });
      this.lists.length > 1
        ? (this.islistdropdowndisabled = false)
        : (this.islistdropdowndisabled = true);
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
