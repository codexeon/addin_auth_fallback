/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { AuthenticationResult } from "@azure/msal-browser";
import {
  AccountContext,
  createLocalUrl,
  ensurePublicClient,
  getAccountFromContext,
  getTokenRequest,
  msalConfig,
} from "./msalcommon";

/* global console, document, Office */

Office.onReady((info) => {
  document.getElementById("msal_js_button").onclick = msalAuth;
  document.getElementById("dialog_api_button").onclick = dialogApiAuth;
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
  Office.context.ui.displayDialogAsync("https://localhost:3000/dialog.html?logout=1");
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
  try {
    const accountContext = await getAccountContext();
    const request = await getTokenRequest(accountContext);
    const pca = await ensurePublicClient();
    let result: AuthenticationResult;
    if (request.account) {
      result = await pca.acquireTokenSilent(request);
    } else {
      if (request.loginHint) {
        try {
          result = await pca.ssoSilent(request);
        } catch {}
      }
      if (!result) {
        result = await pca.acquireTokenPopup(request);
      }
      pca.setActiveAccount(result.account);
    }

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
async function dialogApiAuth() {
  const accountContext = await getAccountContext();
  Office.context.ui.displayDialogAsync(
    createLocalUrl(`dialog.html?accountContext=${encodeURIComponent(JSON.stringify(accountContext))}`),
    (result) => {
      loginDialog = result.value;
      result.value.addEventHandler(Office.EventType.DialogMessageReceived, processDialogMessage);
    }
  );
}
