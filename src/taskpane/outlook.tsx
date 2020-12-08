/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import * as React from "react";
import {renderToStaticMarkup} from "react-dom/server";

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
    const variabile = renderToStaticMarkup(<span>123456</span>);
    Office.context.mailbox.item.body.setSelectedDataAsync(variabile,
        {coercionType: Office.CoercionType.Html},
        (asyncResult) => {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                console.log("Error during insertion", asyncResult.error.message);
            }}
    );

}
