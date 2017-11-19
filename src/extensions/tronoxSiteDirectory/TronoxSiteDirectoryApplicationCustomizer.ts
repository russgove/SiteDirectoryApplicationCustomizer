import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from "@microsoft/sp-application-base";
import pnp from "sp-pnp-js";
import { Dialog } from "@microsoft/sp-dialog";
import { escape, debounce } from "@microsoft/sp-lodash-subset";
import * as strings from "TronoxSiteDirectoryApplicationCustomizerStrings";
const LOG_SOURCE: string = "TronoxSiteDirectoryApplicationCustomizer";
import styles from "./AppCustomizer.module.scss";
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ITronoxSiteDirectoryApplicationCustomizerProperties {
  // this is an example; replace with your own property
  testMessage: string;
  Top: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class TronoxSiteDirectoryApplicationCustomizer
  extends BaseApplicationCustomizer<ITronoxSiteDirectoryApplicationCustomizerProperties> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;
  @override
  public onInit(): Promise<void> {
    debugger;
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    // this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    //this._renderPlaceHolders();
    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
      return this._renderPlaceHolders();
    });
  }
  private _renderPlaceHolders(): Promise<any> {
    debugger;
    let editFormUrl:string,welcomePage: string, title: string, description: string;
    console.log("TronoxSiteDirectoryApplicationCustomizer._renderPlaceHolders()");

    // handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });
      // the extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }
      debugger;
      return pnp.sp.web.lists.getByTitle("Site Information").items.get().then((items) => {
        debugger;
        if (items.length < 1) {
          
          let batch = pnp.sp.createBatch();
          // get the home page, so we can create the skeleton site info
          pnp.sp.web.rootFolder.inBatch(batch).get().then((root) => {
            debugger;
            welcomePage = this.context.pageContext.web.absoluteUrl + "/" + root.WelcomePage;
            
          }).catch((error) => {
            debugger;
            console.log(error);
          });
          // get the home page, site title and description so we can create the skeleton site information
          pnp.sp.web.inBatch(batch).get().then((web => {
            debugger;
            title = web.Title;
            description = web.Description;
          })).catch((error) => {

            debugger;
            console.log(error);
          });
          // get the EditForm for the site info list , so we can link the user back to the list
          pnp.sp.web.lists.getByTitle("Site Information").forms.filter('FormType eq 6').inBatch(batch).get().then((forms => {
            debugger;
            editFormUrl= this.context.pageContext.web.absoluteUrl +"/" + forms[0].url;
           
          })).catch((error) => {

            debugger;
            console.log(error);
          });
          return batch.execute().then((x) => {
            // see http://www.pointtaken.no/blogg/updating-single-and-multi-value-taxonomy-fields-using-pnp-js-core/z
            const termString = '-1;#Global|98587941-8870-4d2a-942f-0beb1982ef66;';
            
            return pnp.sp.web.lists.getByTitle("Site Information").items.add({
              Title: title,
              SiteDescription: description,
              "hf0f9d05de3a4646a1b8810ef201df06":termString
            
            }).then((item) => {
              debugger;
              editFormUrl=editFormUrl+"?Id="+ item.data.Id;
            }).catch((error) => {
              debugger;
            });
          }).catch((error) => {
            debugger;

          });
        }

      }).catch((error) => {
        debugger;
        console.log("list not found");
      });
    }
  }
  private _onDispose(): void {
    console.log("[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.");
  }
}
