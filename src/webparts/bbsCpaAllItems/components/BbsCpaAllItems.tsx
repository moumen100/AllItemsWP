import * as React from "react";
import styles from "./BbsCpaAllItems.module.scss";
import { IBbsCpaAllItemsProps } from "./IBbsCpaAllItemsProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";

import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  DocumentCardLocation,
  DocumentCardType,
  SearchBox
} from "office-ui-fabric-react";
import { ImageFit } from "office-ui-fabric-react/lib/Image";
import { ISize } from "office-ui-fabric-react/lib/Utilities";
import { IAllItems } from "../Models/IAllItems";
import SPHelper from "../Helpers/SPHelpers";
import { DefaultButton, IButtonProps } from "office-ui-fabric-react/lib/Button";
import { Label } from "office-ui-fabric-react/lib/Label";

export default class BbsCpaWebpartOffreEmploi extends React.Component<
  IBbsCpaAllItemsProps,
  { items: IAllItems[],searchQuery:string }
> {
  constructor(props: IBbsCpaAllItemsProps) {
    super(props);

    this.state = {
      items: [],
      searchQuery:""
    };
  }

  

  componentDidUpdate(
    prevProps: Readonly<IBbsCpaAllItemsProps>,
   prevState: Readonly<{ items: IAllItems[] }>,
    prevContext: any
  ): void {
        let SpHelper = new SPHelper(this.props.Context,this.props.WebUrl);

    if (prevProps.listId !== this.props.listId ) 
    {
      SpHelper.getListItems(this.props.listId).then(result=>{
        this.setState({
          items:result.value
        })
      })
    }
  }

  componentDidMount() {
    let SpHelper = new SPHelper(this.props.Context,this.props.WebUrl);
    SpHelper.getListItems(this.props.listId).then(result=>{
      this.setState({
        items:result.value
      })
    })
  }

  private _onRenderGridItem = (
    items: IAllItems,
    finalSize: ISize,
    isCompact: boolean
  ): JSX.Element => {
    const previewProps: IDocumentCardPreviewProps = {
      previewImages: [
        {
          previewImageSrc: items.ImageUrl,
          imageFit: ImageFit.cover,
          height: 130,
        },
      ],
    };
    return (
      <div data-is-focusable={true} role="listitem" aria-label={items.Titre}>
        <DocumentCard
          type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
        >
          <DocumentCardPreview {...previewProps} />
          <label className={styles.ProductTitle}>{items.Titre}
             </label> 

          {/* {!isCompact && <DocumentCardLocation location={items.ShortDescription} />} */}
          <p className={styles["P-Description"]}>
            {items.ShortDescription}
          </p>
        
          <div className={styles.bouttonCenter}>
            <Label>Primary</Label>
            <DefaultButton
              primary={true}
              data-automation-id="test"
              disabled={false}
              checked={false}
              text="Lire plus"
              //onClick={this._alertClicked}
              onClick={(e) => this._alertClicked(items)}
            />
          </div>
        </DocumentCard>
      </div>
    );
  };
  public render(): React.ReactElement<IBbsCpaAllItemsProps> {
    const ListItems = this.state.items
    const renderItems = () => {
      if (ListItems && ListItems.length)
        return ListItems
          .filter((item) => {
            if (this.state.searchQuery.toLowerCase() === "") return item;
            else if (
              item.Titre.toLowerCase().indexOf(
                this.state.searchQuery.toLowerCase()
              ) > -1 ||
              item.ShortDescription.toLowerCase().indexOf(
                this.state.searchQuery.toLowerCase()
              ) > -1
            )
              return item;
          })
          
    };
    return (
      <div>
         {ListItems ? (<div>
        
        <Label className={styles.LinkTitleStyle}>
            {this.props.description}
          </Label>
      <div className={styles.container_block}>
            <div className={styles.element1}>
              <SearchBox
                placeholder="Rechercher ..."
                onEscape={(ev) => {
                  //console.log("Custom onEscape Called");
                }}
                onClear={(ev) => {
                  // console.log("Custom onClear Called");
                }}
                onChange={(newValue) =>
                  this.setState({ searchQuery: newValue })
                }
                onSearch={(newValue) =>
                  this.setState({ searchQuery: newValue })
                }
              />

              <br />
            </div>
            </div>
      <GridLayout
        items={renderItems()}
        onRenderGridItem={(item: IAllItems, finalSize: ISize, isCompact: boolean) =>
          this._onRenderGridItem(item, finalSize, isCompact)
        }
      />
    </div>) : (
          <p>
            Aucune donnée n'est disponible, veuillez alimenter votre sources de
            données
          </p>
        ) }
      </div>
     
      
    );
  }
  private _alertClicked(items: IAllItems): void {
    window.open(items.ProductUrl, "_blank");
  }
}
