import * as React from "react";
import styles from "./BbsCpaAllItems.module.scss";
import { IBbsCpaAllItemsProps } from "./IBbsCpaAllItemsProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from "office-ui-fabric-react/lib/ChoiceGroup";
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
import { Link, Image, List } from "office-ui-fabric-react";
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
    prevProps: Readonly<IBbsCpaAllItemsProps>
  ): void {
    if (
      prevProps != this.props 
    ) {
      this.setState({
        items:this.props.listitem
      })
    }
  }

  componentWillMount(): void {
    this.setState({
      items:this.props.listitem
    })  }

  componentDidMount() {
    this.setState({
      items:this.props.listitem
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
          <DocumentCardTitle title={items.Titre} shouldTruncate={true} />

          {/* {!isCompact && <DocumentCardLocation location={items.ShortDescription} />} */}
          <p className={styles["P-Description"]}>
            {items.ShortDescription.substring(0, 160) + "... "}{" "}
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
    const renderItems = () => {
      if (this.state.items && this.state.items.length)
        return this.state.items
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
      </div>
    );
  }
  private _alertClicked(items: IAllItems): void {
    window.open(items.ProductUrl, "_blank");
  }
}
