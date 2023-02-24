/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { Configuration, OpenAIApi } from "openai";

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

function getSelectedText(): Promise<any> {
  return new Office.Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
        const configuration = new Configuration({
          apiKey: "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",
        });
        const openai = new OpenAIApi(configuration);
        const response = await openai.createCompletion({
          model: "text-davinci-003",
          prompt: "Turn the following text into a professional business mail: " + asyncResult.value,
          temperature: 0.7,
          max_tokens: 300,
        });

        resolve(response.data.choices[0].text);
      });
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
// function action(event: Office.AddinCommands.Event) {
//   const message: Office.NotificationMessageDetails = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true,
//   };

//   // Show a notification message
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

//   // Be sure to indicate when the add-in command function is complete
//   event.completed();
// }
function action(event: Office.AddinCommands.Event) {
  getSelectedText().then(function (selectedText) {
    Office.context.mailbox.item.setSelectedDataAsync(selectedText, { coercionType: Office.CoercionType.Text });
    event.completed();
  });
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
