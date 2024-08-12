/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { createNestablePublicClientApplication } from "@azure/msal-browser";

let pca = undefined;
Office.onReady(async () => {
  // Initialize the public client application
  pca = await createNestablePublicClientApplication({
    auth: {
      clientId: "605f8396-522e-4d3c-a83d-829fd2fcf47e",
      authority: "https://login.microsoftonline.com/common",
    },
  });
});

async function getUserName() {
  // Specify minimum scopes needed for the access token.
  const tokenRequest = {
    scopes: ["Files.Read", "User.Read", "openid", "profile"],
  };
  let accessToken = null;

  // TODO 1: Call acquireTokenSilent.
  try {
    //console.log("Trying to acquire token silently...");
    const userAccount = await pca.acquireTokenSilent(tokenRequest);
    // console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (error) {
    //console.log(`Unable to acquire token silently: ${error}`);
  }
  // TODO 2: Call acquireTokenPopup.
  if (accessToken === null) {
    // Acquire token silent failure. Send an interactive request via popup.
    try {
      // console.log("Trying to acquire token interactively...");
      const userAccount = await pca.acquireTokenPopup(tokenRequest);
      // console.log("Acquired token interactively.");
      accessToken = userAccount.accessToken;
    } catch (popupError) {
      // Acquire token interactive failure.
      //  console.log(`Unable to acquire token interactively: ${popupError}`);
    }
  }
  // TODO 3: Log error if token still null.
  // Log error if both silent and popup requests failed.
  if (accessToken === null) {
    // console.error(`Unable to acquire access token.`);
    return;
  }
  // TODO 4: Call the Microsoft Graph API.
  // Call the Microsoft Graph API with the access token.
  const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root/children?$select=name&$top=10`, {
    headers: { Authorization: accessToken },
  });

  if (response.ok) {
    // Write file names to the console.
    const data = await response.json();
    const names = data.value.map((item) => item.name);

    // Be sure the taskpane.html has an element with Id = item-subject.
    // const label = document.getElementById("item-subject");

    // Write file names to task pane and the console.
    const nameText = names.join(", ");
    //if (label) label.textContent = nameText;
    // console.log(nameText);
    return nameText;
  } else {
    const errorText = await response.text();
    // console.error("Microsoft Graph call failed - error text: " + errorText);
  }
}

function onNewMessageComposeHandler(event) {
  setSubject(event);
}
function onNewAppointmentComposeHandler(event) {
  setSubject(event);
}
async function setSubject(event) {
  let name = await getUserName();

  Office.context.mailbox.item.subject.setAsync(
    "Set from src by an event-based add-in!" + name,
    {
      asyncContext: event,
    },
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
      }

      // Call event.completed() to signal to the Outlook client that the add-in has completed processing the event.
      asyncResult.asyncContext.completed();
    }
  );
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
}
