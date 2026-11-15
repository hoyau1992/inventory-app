(function (global) {
  "use strict";

  // Application (client) ID from the Entra app registration.
  // Leave as a99f8e1d-53de-4126-bc38-e334972e4bd2 until the parent fills in the real ID.
  // This value is public in a browser app. There is no client secret.
  var clientId = "a99f8e1d-53de-4126-bc38-e334972e4bd2";

  global.INVENTORY_CONFIG = {
    msal: {
      clientId: clientId,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: "https://hoyau1992.github.io/inventory-app/"
    },
    scopes: ["Files.ReadWrite.AppFolder", "User.Read"],
    syncFileName: "inventory-sync.json"
  };
})(typeof window !== "undefined" ? window : globalThis);
