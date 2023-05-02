import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-users";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";

var _sp: SPFI = null;

const getSP = (context?: WebPartContext): SPFI => {
  _sp = spfi().using(SPFx(context));
  return _sp;
};

export default getSP;
