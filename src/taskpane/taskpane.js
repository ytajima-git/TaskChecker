/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    var app_body = document.getElementById("app-body");
    app_body.children[0].innerHTML = "<b>Get Message</b> <br/>";
    var item = Office.context.mailbox.item;
    item.body.getAsync(
      "text",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        const message = result.value;
        const TinySegmenter = require('tiny-segmenter');
        const tinySegmenter = new TinySegmenter();

        // const tfjs = require('@tensorflow/tfjs');
        // require('@tensorflow/tfjs-node');

        const segments = tinySegmenter.segment(message);
        document.getElementById("item-subject").innerHTML = "<b>Message:</b> <br/>" + segments;
      });
  };
});
