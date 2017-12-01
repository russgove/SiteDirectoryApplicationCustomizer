import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from "@microsoft/sp-application-base";
import pnp from "sp-pnp-js";
import { Site } from "sp-pnp-js";
import { Dialog } from "@microsoft/sp-dialog";
import { escape, debounce } from "@microsoft/sp-lodash-subset";
import * as strings from "TronoxSiteDirectoryApplicationCustomizerStrings";
const LOG_SOURCE: string = "TronoxSiteDirectoryApplicationCustomizer";
import styles from "./AppCustomizer.module.scss";
import SPPermission from "@microsoft/sp-page-context/lib/SPPermission";
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
require("sp-taxonomy");

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

  @override
  public onInit(): Promise<void> {
    // Need to be admin in order to remove the customizer - if not skip doing the work
    // For Group sites, the owners will be site admins
    let isSiteAdmin = this.context.pageContext.legacyPageContext.isSiteAdmin;
    if (isSiteAdmin) {
      this.DoWork();
    }
    return Promise.resolve();
  }
  private async DoWork() {
    pnp.setup({
      spfxContext: this.context
    });
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
    }
    // use await if you want to block the dialog before continue
    //await Dialog.alert(data);
    let editFormUrl: string, welcomePage: string, title: string, description: string;
    return pnp.sp.web.lists.getByTitle("Site Information").items.get()
      .then((items) => {
        debugger;
        console.log("there are " + items.length + " items in the site info list");
        if (items.length < 1) { // create the item in the site info list
          let batch = pnp.sp.createBatch();
          // get the home page, site title and description so we can create the skeleton site information
          pnp.sp.web.rootFolder.inBatch(batch).get().then((root) => {
            welcomePage = this.context.pageContext.web.absoluteUrl + "/" + root.WelcomePage;
          }).catch((error) => {
            debugger;
            console.log(error);
          });
          pnp.sp.web.inBatch(batch).get().then((web => {
            title = web.Title;
            description = web.Description;
          })).catch((error) => {
            debugger;
            console.log(error);
          });
          // get the EditForm for the site info list , so we can link the user back to the list
          pnp.sp.web.lists.getByTitle("Site Information").forms.filter('FormType eq 6').inBatch(batch).get().then((forms => {
            editFormUrl = "https://" + window.location.hostname + "/" + forms[0].ServerRelativeUrl;
          })).catch((error) => {
            debugger;
            console.log(error);
          });
          return batch.execute().then((x) => {
            const updateObject = {
              Title: title,
              SiteDescription: description,
              SiteLink: {
                '__metadata': { 'type': 'SP.FieldUrlValue' },
                'Description': title,
                'Url': welcomePage
              }
            }
            return pnp.sp.web.lists.getByTitle("Site Information").items.add(updateObject).then((item) => {
              debugger;
              // these fiedlds needed in jsom callback
              const itemId = item.data.Id;
              const mystyles = styles;
              const myCustomizer = this;

              //need to switch to jsom to work with managed metadata
              const context: SP.ClientContext = new SP.ClientContext(this.context.pageContext.site.serverRelativeUrl);
              var list = context.get_web().get_lists().getByTitle("Site Information");
              var items = list.getItems(SP.CamlQuery.createAllItemsQuery());
              context.load(items);
              var field = list.get_fields().getByInternalNameOrTitle("SiteLocation");
              let txField: SP.Taxonomy.TaxonomyField = context.castTo(field, SP.Taxonomy.TaxonomyField) as SP.Taxonomy.TaxonomyField;
              context.load(txField);
              context.executeQueryAsync(function (sender, args) {
                //1. Prepare TaxonomyFieldValueCollection object
                var terms = new Array();
                terms.push("-1;#Global|98587941-8870-4d2a-942f-0beb1982ef66");
                var termValueString = terms.join(";#");
                var termValues = new SP.Taxonomy.TaxonomyFieldValueCollection(context, termValueString, txField);
                //2. Update multi-valued taxonomy field
                var item = items.getItemAtIndex(0)
                txField.setFieldValueByValueCollection(item, termValues);
                item.update();
                context.executeQueryAsync(function () {
                  debugger;
                  editFormUrl = editFormUrl + "?Id=" + itemId + "&SourceUrl=" + myCustomizer.context.pageContext.web.absoluteUrl;
                  let message = "A 'Site Information' list has been created in this site which will be used to display the site in the Tronox Site Directory, and default site information has been added. Please click <a href='" + editFormUrl + "'>here</a> to verify and complete the site information";
                  myCustomizer._topPlaceholder.domElement.innerHTML = `
                      <div class="${styles.app}">
                        <div class="ms-bgColor-themeDark ms-fontColor-white ${mystyles.top}">
                         <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${message}
                        </div>
                     </div>`;
                  myCustomizer.removeCustomizer();
                });
              });
            }).catch((error) => {
              console.log(error.data.responseBody["odata.error"].message.value);
              debugger;
            });
          }).catch((error) => {
            debugger;
          });
        }
      })
      .catch((error) => {
        debugger;
        console.log("Site Information list not found");
      });
  }
  private async removeCustomizer() {
    debugger;
    // Remove custom action from current sute
    return async function () {
      let site = new Site(this.context.pageContext.site.absoluteUrl);
      let customActions = await site.userCustomActions.get(); // if installed as web scope, change this line
      for (let i = 0; i < customActions.length; i++) {
        var instance = customActions[i];
        if (instance.ClientSideComponentId === this.componentId) {
          await site.userCustomActions.getById(instance.Id).delete();
          console.log("Extension removed");
          // reload the page once done if needed
          window.location.href = window.location.href;
          break;
        }
      }
    }
  }
  // private _renderPlaceHolders(): Promise<any> {
  //   debugger;
  //   let editFormUrl: string, welcomePage: string, title: string, description: string;
  //   console.log("TronoxSiteDirectoryApplicationCustomizer._renderPlaceHolders()");
  //   // handling the top placeholder
  //   if (!this._topPlaceholder) {
  //     this._topPlaceholder =
  //       this.context.placeholderProvider.tryCreateContent(
  //         PlaceholderName.Top,
  //         { onDispose: this._onDispose });
  //     // the extension should not assume that the expected placeholder is available.
  //     if (!this._topPlaceholder) {
  //       console.error("The expected placeholder (Top) was not found.");
  //       return;
  //     }
  //     return pnp.sp.web.lists.getByTitle("Site Information").items.select("*,j38a792f394e4b1b87278c688d74df04,hf0f9d05de3a4646a1b8810ef201df06").get()
  //       .then((items) => {
  //         debugger;
  //         console.log("there are " + items.length + "items in the site info list");
  //         if (items.length < 1) { // create the item in the site info list
  //           let batch = pnp.sp.createBatch();
  //           // get the home page, so we can create the skeleton site info
  //           pnp.sp.web.rootFolder.inBatch(batch).get().then((root) => {
  //             welcomePage = this.context.pageContext.web.absoluteUrl + "/" + root.WelcomePage;
  //           }).catch((error) => {
  //             debugger;
  //             console.log(error);
  //           });
  //           // get the home page, site title and description so we can create the skeleton site information
  //           pnp.sp.web.inBatch(batch).get().then((web => {

  //             title = web.Title;
  //             description = web.Description;
  //           })).catch((error) => {

  //             debugger;
  //             console.log(error);
  //           });
  //           // get the EditForm for the site info list , so we can link the user back to the list
  //           pnp.sp.web.lists.getByTitle("Site Information").forms.filter('FormType eq 6').inBatch(batch).get().then((forms => {

  //             editFormUrl = "https://" + window.location.hostname + "/" + forms[0].ServerRelativeUrl;

  //           })).catch((error) => {

  //             debugger;
  //             console.log(error);
  //           });
  //           return batch.execute().then((x) => {
  //             // see http://www.pointtaken.no/blogg/updating-single-and-multi-value-taxonomy-fields-using-pnp-js-core/
  //             // site department hidden note field internal name is j38a792f394e4b1b87278c688d74df04
  //             // site location hidden note field internal name is e2a303c6-777d-4806-b63d-5815aeaa1d5c

  //             // const siteLocationTermString = '-1;#Global|98587941-8870-4d2a-942f-0beb1982ef66;';// global
  //             // const siteDepartmentString = '-1;#All|6595e644-60a7-42fb-955c-c31ecafd4431;';// global
  //             const siteLocationTermString = "-1;#Global|98587941-8870-4d2a-942f-0beb1982ef66;#-1;#Americas|7af02390-41dd-471d-b8bb-4d8559abbb78;";// global
  //             const siteDepartmentString = "-1;#All|6595e644-60a7-42fb-955c-c31ecafd4431;#-1;#Finance|76ca6c85-e8c3-456b-a23b-f28a65ef002a;"// all 

  //             const updateObject = {
  //               Title: title,
  //               SiteDescription: description,
  //               SiteLink: {
  //                 '__metadata': { 'type': 'SP.FieldUrlValue' },
  //                 'Description': title,
  //                 'Url': welcomePage
  //               },
  //               "j38a792f394e4b1b87278c688d74df04": siteDepartmentString, // hf0... is the internal name 
  //               "hf0f9d05de3a4646a1b8810ef201df06": siteLocationTermString


  //             }
  //             return pnp.sp.web.lists.getByTitle("Site Information").items.add(updateObject, "SP.Data.SiteInformationListItem").then((item) => {
  //               debugger;
  //               // these fiedlds needed in jsom callback
  //               const itemId = item.data.Id;
  //               const mystyles = styles;
  //               const myCustomizer = this;


  //               const context: SP.ClientContext = new SP.ClientContext(this.context.pageContext.site.serverRelativeUrl);
  //               var list = context.get_web().get_lists().getByTitle("Site Information");
  //               var items = list.getItems(SP.CamlQuery.createAllItemsQuery());
  //               context.load(items);
  //               var field = list.get_fields().getByInternalNameOrTitle("SiteLocation");
  //               let txField: SP.Taxonomy.TaxonomyField = context.castTo(field, SP.Taxonomy.TaxonomyField) as SP.Taxonomy.TaxonomyField;
  //               context.load(txField);
  //               context.executeQueryAsync(function (sender, args) {
  //                 //1. Prepare TaxonomyFieldValueCollection object
  //                 var terms = new Array();
  //                 terms.push("-1;#Global|98587941-8870-4d2a-942f-0beb1982ef66");
  //                 var termValueString = terms.join(";#");
  //                 var termValues = new SP.Taxonomy.TaxonomyFieldValueCollection(context, termValueString, txField);
  //                 //2. Update multi-valued taxonomy field
  //                 var item = items.getItemAtIndex(0)
  //                 txField.setFieldValueByValueCollection(item, termValues);
  //                 item.update();
  //                 context.executeQueryAsync(function () {
  //                   debugger;
  //                   editFormUrl = editFormUrl + "?Id=" + itemId + "&SourceUrl=" + myCustomizer.context.pageContext.web.absoluteUrl;
  //                   let message = "The 'Site Information' list has been created in your site and a default entry has been added. Please click <a href='" + editFormUrl + "'>here</a> to verify and complete the site information";
  //                   myCustomizer._topPlaceholder.domElement.innerHTML = `
  //                 <div class="${styles.app}">
  //                   <div class="ms-bgColor-themeDark ms-fontColor-white ${mystyles.top}">
  //                     <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${message}
                     
  //                   </div>
                    
  //                 </div>`;
  //                   myCustomizer.removeCustomizer();
  //                 });
  //               });



  //             }).catch((error) => {
  //               console.log(error.data.responseBody["odata.error"].message.value);
  //               debugger;
  //             });
  //           }).catch((error) => {
  //             debugger;

  //           });
  //         } else { // remind user to complete the entry

  //           const item = items[0];
  //           let needsUpdating = false;
  //           if (item["SiteDepartment"].length < 1) {
  //             needsUpdating = true;
  //             console.log("Site department is missing from th esite infoemation list");
  //           }

  //           if (!item["SiteLocation"].length) {
  //             needsUpdating = true;
  //           }
  //           if (needsUpdating) {
  //             // get the EditForm for the site info list , so we can link the user back to the list
  //             return pnp.sp.web.lists.getByTitle("Site Information").forms.filter('FormType eq 6').get().then((forms => {

  //               editFormUrl = "https://" + window.location.hostname + "/" + forms[0].ServerRelativeUrl + "?ID=" + items[0].Id;
  //               let message = "Please complete the information in the 'Site Information' list <a href='" + editFormUrl + "'>here</a> so that your site can be listed in the Tronox Site Directory.";
  //               this._topPlaceholder.domElement.innerHTML = `
  //               <div class="${styles.app}">
  //                 <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">
  //                   <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${message}
  //                   <button onClick=${ this.removeCustomizer}>Dont't show this again! </button>
  //                 </div>
  //                 </div>

  //               </div>`;

  //             })).catch((error) => {

  //               debugger;
  //               console.log(error);
  //             });
  //           }


  //         }

  //       })
  //       .catch((error) => {
  //         debugger;
  //         console.log("list not found");
  //       });
  //   }
  // }
  private _onDispose(): void {
    console.log("[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.");
  }
}
