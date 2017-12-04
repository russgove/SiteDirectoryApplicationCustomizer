import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from "@microsoft/sp-application-base";
import pnp from "sp-pnp-js";
import { Site, UserCustomActionAddResult } from "sp-pnp-js";
import { Dialog } from "@microsoft/sp-dialog";
import { escape, debounce } from "@microsoft/sp-lodash-subset";
import * as strings from "TronoxSiteDirectoryApplicationCustomizerStrings";
const LOG_SOURCE: string = "TronoxSiteDirectoryApplicationCustomizer";
import styles from "./AppCustomizer.module.scss";
import SPPermission from "@microsoft/sp-page-context/lib/SPPermission";
import { UserCustomAction, UserCustomActions } from "sp-pnp-js/lib/sharepoint/usercustomactions";
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
    console.log("User is Site Admin:" + isSiteAdmin);
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

              // these fiedlds needed in jsom callback
              const itemId = item.data.Id;
              const mystyles = styles;
              const myCustomizer = this;

              //need to switch to jsom to work with managed metadata
              const context: SP.ClientContext = new SP.ClientContext(this.context.pageContext.site.serverRelativeUrl);
              var list = context.get_web().get_lists().getByTitle("Site Information");
              var items = list.getItems(SP.CamlQuery.createAllItemsQuery());
              context.load(items);
              // Site Location
              var siteLocationField = list.get_fields().getByInternalNameOrTitle("SiteLocation");
              let siteLocationTxField: SP.Taxonomy.TaxonomyField = context.castTo(siteLocationField, SP.Taxonomy.TaxonomyField) as SP.Taxonomy.TaxonomyField;
              context.load(siteLocationTxField);
              // Site Department
              var siteDepartmentField = list.get_fields().getByInternalNameOrTitle("SiteDepartment");
              let siteDepartmentTxField: SP.Taxonomy.TaxonomyField = context.castTo(siteDepartmentField, SP.Taxonomy.TaxonomyField) as SP.Taxonomy.TaxonomyField;
              context.load(siteDepartmentTxField);

              context.executeQueryAsync(function (sender: any, args: any) {
                var item = items.getItemAtIndex(0)
                //site Department
                var siteLocationTermValueString = "-1;#Global|98587941-8870-4d2a-942f-0beb1982ef66";
                var siteLocationTermValues = new SP.Taxonomy.TaxonomyFieldValueCollection(context, siteLocationTermValueString, siteLocationTxField);
                siteLocationTxField.setFieldValueByValueCollection(item, siteLocationTermValues);
                //site Department
                const siteDepartmentTermValueString = "-1;#All|6595e644-60a7-42fb-955c-c31ecafd4431";
                var siteDepartmentTermValues = new SP.Taxonomy.TaxonomyFieldValueCollection(context, siteDepartmentTermValueString, siteDepartmentTxField);
                siteDepartmentTxField.setFieldValueByValueCollection(item, siteDepartmentTermValues);

                item.update();
                context.executeQueryAsync(function () {

                  editFormUrl = editFormUrl + "?ID=" + itemId + "&SourceUrl=" + myCustomizer.context.pageContext.web.absoluteUrl;
                  let message = "A list titled 'Site Information' has been created in this site which will be used to display the site in the Tronox Site Directory, and default site information has been added.<br  /> Please click <a href='" + editFormUrl + "'>here</a> to verify and complete the site information. <br />You can edit the item in the Site Information list to update your listing in sthe Tronox Site Directory at any time (It may take a few hours for your changes to be refelected in the directory).";
                  myCustomizer._topPlaceholder.domElement.innerHTML = `
                      <div class="${styles.app}">
                        <div class="ms-bgColor-themeDark ms-fontColor-white ${mystyles.top}">
                         ${message}
                        </div>
                     </div>`;
                  myCustomizer.removeCustomizer();
                }, function (sender: any, args: SP.ClientRequestFailedEventArgs) {
                  debugger;
                  console.log(args.get_message());
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
        else {
          console.log("There is an item in the list, removing custom action");
          debugger;
          this.removeCustomizer();
        }
      })
      .catch((error) => {
        debugger;
        console.log("Site Information list not found");
      });
  }
  private async removeCustomizer() {
    console.log("in  removeCustomizer")
    try {
      // Remove custom action from current sute
      let site = new Site(this.context.pageContext.site.absoluteUrl);
      site.rootWeb.userCustomActions.get().then(customActions => { // if installed as web scope, change this line
        for (let i = 0; i < customActions.length; i++) {
          var instance = customActions[i];
          if (instance.ClientSideComponentId === this.componentId) {
            debugger;
            site.rootWeb.userCustomActions.getById(instance.Id).delete().then((ss) => {
              console.log("Extension removed");
            }).catch((error) => {
              debugger;
              console.log("an error occurred removing the userCustomAction");
              console.log(error);
            });
            // reload the page once done if needed
            //window.location.href = window.location.href;
            break;
          }
        }
      });
    }
    catch (e) {
      debugger;
      console.log("an error occurred in removeCustomizer")
    }
  }

  private _onDispose(): void {
    console.log("[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.");
  }
}
