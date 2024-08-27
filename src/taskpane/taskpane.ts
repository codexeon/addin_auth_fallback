/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { AuthenticationResult } from "@azure/msal-browser";
import { AccountContext, ensurePublicClient, getTokenRequest } from "./msalcommon";
import { createLocalUrl } from "./util";

/* global console, document, Office, window */

Office.onReady((info) => {
  document.getElementById("msal_js_button").onclick = msalAuth;
  document.getElementById("dialog_api_button").onclick = () => dialogApiAuth(false);
  document.getElementById("dialog_api_button_ie").onclick = () => dialogApiAuth(true);
  document.getElementById("sign_out_button").onclick = signOut;
  document.getElementById("sign_out_dialog_button").onclick = signOutDialogAPi;

  checkSignedIn();
});

function output(text: string) {
  document.getElementById("output").innerText = text;
}

let _authContext: AccountContext | null = null;

async function getAccountContext(): Promise<AccountContext | null> {
  if (!_authContext) {
    try {
      const authContext = await (Office.auth as any).getAuthContext();
      _authContext = {
        loginHint: authContext.loginHint,
        tenantId: authContext.tenantId,
        localAccountId: authContext.userObjectId,
      };
    } catch {
      _authContext = {};
    }
  }
  return _authContext;
}

async function checkSignedIn() {
  const pca = await ensurePublicClient();
  output(pca.getAllAccounts().length.toString());
}

function signOutDialogAPi() {
  Office.context.ui.displayDialogAsync(createLocalUrl("dialog.html?logout=1"));
}

async function signOut() {
  try {
    const pca = await ensurePublicClient();
    await pca.logoutPopup();
    output("Signed out");
  } catch (ex) {
    output(ex);
  }
}

async function msalAuth() {
  sessionStorage.clear();

  try {
    const accountContext = await getAccountContext();
    const request = await getTokenRequest(accountContext);
    const pca = await ensurePublicClient();
    let result: AuthenticationResult;
    try {
      if (request.account) {
        result = await pca.acquireTokenSilent(request);
      } else {
        if (request.loginHint) {
          result = await pca.ssoSilent(request);
        }
      }
    } catch {}

    if (!result) {
      result = await pca.acquireTokenPopup(request);
    }
    pca.setActiveAccount(result.account);
    output(result.accessToken);
  } catch (ex) {
    output(ex);
  }
}

function processDialogMessage(arg: { message: string; origin: string | undefined }) {
  const parsedMessage = JSON.parse(arg.message);
  output(parsedMessage.token);
  loginDialog.close();
}

let loginDialog: Office.Dialog;
async function dialogApiAuth(isInternetExplorer?: boolean) {
  const accountContext = await getAccountContext();
  Office.context.ui.displayDialogAsync(
    createLocalUrl(
      `${isInternetExplorer ? "dialogie.html" : "dialog.html"}?accountContext=${encodeURIComponent(JSON.stringify(accountContext))}`
    ),
    (result) => {
      loginDialog = result.value;
      result.value.addEventHandler(Office.EventType.DialogMessageReceived, processDialogMessage);
    }
  );
}
