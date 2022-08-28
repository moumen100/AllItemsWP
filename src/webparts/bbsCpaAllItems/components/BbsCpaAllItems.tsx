import * as React from "react";
import styles from "./BbsCpaAllItems.module.scss";
import { IBbsCpaAllItemsProps } from "./IBbsCpaAllItemsProps";
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";

import {
  DocumentCard,
  DocumentCardPreview,
  IDocumentCardPreviewProps,
  DocumentCardType,
  SearchBox
} from "office-ui-fabric-react";
import { ImageFit } from "office-ui-fabric-react/lib/Image";
import { ISize } from "office-ui-fabric-react/lib/Utilities";
import { IAllItems } from "../Models/IAllItems";
import SPHelper from "../Helpers/SPHelpers";
import { DefaultButton } from "office-ui-fabric-react/lib/Button";
import { Label } from "office-ui-fabric-react/lib/Label";

export default class BbsCpaWebpartOffreEmploi extends React.Component<
  IBbsCpaAllItemsProps,
  { items: IAllItems[],searchQuery:string }
> {
 /**
  * constructor afin d'initialiser state de react 
  * @param props Props react en paramètre 
  */
  constructor(props: IBbsCpaAllItemsProps) {
    super(props);
    /**
     * initiliaser list des items 
     * initiliser mot clés de recherche  
     */    
    this.state = {
      items: [],
      searchQuery:""
    };
  }

    /**
     * évènement  utilisé par react se déclanche au changement des props 
     * @param prevProps // ancien valeur des props 
     * @param prevState // ancien valeur des state 
     */

  componentDidUpdate(
    prevProps: Readonly<IBbsCpaAllItemsProps>,
   prevState: Readonly<{ items: IAllItems[] }>,
  ): void {
        let SpHelper = new SPHelper(this.props.Context,this.props.WebUrl); // initiliser l'objet helper afin d'utiliser les méthodes de la class sphelper 

    if (prevProps.listId !== this.props.listId ) //condition pour tester les props ont été changé
    {
          // fait appel à la méthode getListItems définit dans la class SpHelper Afin de récupéré la liste de sproduits 
      SpHelper.getListItems(this.props.listId).then(result=>{ 
        this.setState({
          items:result.value // modifier le state de react avec les valeurs de la liste des produits 
        })
      })
    }
  }
  /**
   * évènement react déclanche une fois lorsque le component se montre  la première fois 
   */
  componentDidMount() {
    let SpHelper = new SPHelper(this.props.Context,this.props.WebUrl); // initiliser l'objer de la class SpHemper afin d'utiliser les méthode de cztte class 
    // fait appel à la méthode getListItems définit dans la class SpHelper Afin de récupéré la liste de sproduits 
    SpHelper.getListItems(this.props.listId).then(result=>{
      this.setState({
        items:result.value // modifier le state de react avec les valeurs de la liste des produits 
      })
    })
  }
  /**
   * Cette méthode permet d'afficher les  elments de la liste de tout les produits sous forme d'une grid ( cards) 
   * EN utilisants le coposant DocumentCard définit dans la library react fabric ui ( fluent Ui) 
   * @param item //represente un item (un produit)  de la liste des produit 
   * @param finalSize 
   * @param isCompact //type d'affichage de la grid Items 
   * @returns 
   */
  private _onRenderGridItem = (
    item: IAllItems,
    finalSize: ISize,
    isCompact: boolean
  ): JSX.Element => {
    const previewProps: IDocumentCardPreviewProps = {
      previewImages: [
        {
          previewImageSrc: item.ImageUrl,
          imageFit: ImageFit.cover,
          height: 130,
        },
      ],
    };
    return (
      <div data-is-focusable={true} role="listitem" aria-label={item.Titre}>
        <DocumentCard
          type={isCompact ? DocumentCardType.compact : DocumentCardType.normal}
        >
          <DocumentCardPreview {...previewProps} />
          <label className={styles.ProductTitle}>{item.Titre}
             </label> 

          {/* {!isCompact && <DocumentCardLocation location={items.ShortDescription} />} */}
          <p className={styles["P-Description"]}>
            {item.ShortDescription}
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
              onClick={(e) => this._alertClicked(item)}
            />
          </div>
        </DocumentCard>
      </div>
    );
  };
  /**
   * Méthode permet d'afficher le component de la webpart 
   * @returns contenu de comonent react
   */
  public render(): React.ReactElement<IBbsCpaAllItemsProps> {
    /**
     * Tout d'abord avant d'afficher  ont doit filtrer les éléments à afficher selon le text de la recherche des produits 
     * Si la bar de recherche est vide on affiche toute la liste 
     */
    const ListItems = this.state.items // sauvegarde la liste filtere 
    /**
     * Fonction constante qui retourne une liste filtré selons les mots clés recherché
     * @returns Liste filtré 
     */
    const renderItems = () => {
      if (ListItems && ListItems.length)
        return ListItems
          .filter((item) => {
            if (this.state.searchQuery.toLowerCase() === "") return item; // cas text champs vide retourné tout les items
            else if (
              item.Titre.toLowerCase().indexOf( //recherche dans  le titre 
                this.state.searchQuery.toLowerCase()
              ) > -1 ||
              item.ShortDescription.toLowerCase().indexOf( //recherche dans la description
              this.state.searchQuery.toLowerCase()
              ) > -1
            )
              return item; //retourner les items filtré 
          })
          
    };
    return (
      <div>
         {ListItems ? (<div>
        {/* Partie d'afficher la bar de recherche */}
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
            {/* partie pour afficher le grid 
              * On utilise le component GridLayout définit dans la library @pnp/spfx-controls-react/
            */}
      <GridLayout
      //  récupérer les items filtré 
        items={renderItems()} 
        onRenderGridItem={(item: IAllItems, finalSize: ISize, isCompact: boolean) =>
          this._onRenderGridItem(item, finalSize, isCompact) //appel à la méthode _onRenderGridItem pour afficher la Grid
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
  /**
   *    * Cette méthode permet de déclanché lors de l'évenment click sur un elment de la liste des produits 
   * @param items item d'un élément de la grid 
   */
  private _alertClicked(item: IAllItems): void {
    window.open(item.ProductUrl, "_blank"); // accéder au lien vers la page de produit spécifié
  }
}
