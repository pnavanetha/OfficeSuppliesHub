import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


import { SPFx } from "@pnp/sp/presets/all";

export default class YourWebPart {
  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    return Promise.resolve();
  }
}

async getItems() {
    const items = await sp.web.lists
    .getByTitle("ListName")
    .items.get();

    console.log(items);
    }