import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from "@microsoft/sp-application-base";
import pnp from "sp-pnp-js";
import { Dialog } from "@microsoft/sp-dialog";
import { escape } from "@microsoft/sp-lodash-subset";
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
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
    this._renderPlaceHolders();
    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
    });
  }
  private _renderPlaceHolders(): void {


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
      pnp.sp.web.lists.getByTitle("Site Information").items.get().then((items) => {
        debugger;
        if (items.length < 1) {
          if (this._topPlaceholder.domElement) {
            this._topPlaceholder.domElement.innerHTML = `
            <div class="${styles.app}">
              <div class="ms-bgColor-themeDark ms-fontColor-white ${styles.top}">
                <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape("do it dude!")}
              </div>
            </div>`;
          }
        }
      });

    }
  }
  private _onDispose(): void {
    console.log("[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.");
  }
}
