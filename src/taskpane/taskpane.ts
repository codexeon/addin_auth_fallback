/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { BrowserAuthError } from "@azure/msal-browser";
import { AccountContext, ensurePublicClient, getTokenRequest } from "./msalcommon";
import { createLocalUrl } from "./util";
import { AuthenticationResult } from "@azure/msal-browser-v2";
import { clientId } from "./msalconfig";

/* global console, document, Office, window */

Office.onReady(() => {
  document.getElementById("msal_js_button").onclick = authenticate;
  document.getElementById("dialog_api_button").onclick = () => outputResultForDialogApi(false);
  document.getElementById("dialog_api_button_ie").onclick = () => outputResultForDialogApi(true);
  document.getElementById("admin_consent_button").onclick = adminConsent;
  document.getElementById("sign_out_button").onclick = signOut;

  checkSignedIn();
});

function output(text: string) {
  document.getElementById("output").value = text;
}

let _authContext: AccountContext | null = null;

function adminConsent() {
  window.location.href = `https://login.microsoftonline.com/organizations/v2.0/adminconsent?client_id=${clientId}&scope=.default&redirect_uri=https://testnaafallback:3000/auth.html`;
}
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

async function signOut() {
  try {
    const pca = await ensurePublicClient();
    await pca.logoutPopup();
    output("Signed out");
  } catch (ex) {
    output(ex);
  }
}

function showAuthPrompt(): Promise<boolean> {
  return Promise.resolve(true);
}

async function authenticate() {
  try {
    const accountContext = await getAccountContext();
    const request = await getTokenRequest(accountContext);
    const pca = await ensurePublicClient();
    let accessToken: string | null = null;
    let authResult: AuthenticationResult | null = null;
    try {
      if (request.account) {
        authResult = await pca.acquireTokenSilent(request);
      } else {
        if (request.loginHint) {
          authResult = await pca.ssoSilent(request);
        }
      }
    } catch {}

    accessToken = authResult?.accessToken;
    if (!accessToken) {
      const userClickedSignIn = await showAuthPrompt();
      if (userClickedSignIn) {
        try {
          authResult = await pca.acquireTokenPopup(request);
          if (authResult.account) {
            // For fallback flow, set the active account so that next session can be silent SSO
            pca.setActiveAccount(authResult.account);
          }
          accessToken = authResult.accessToken;
        } catch (ex) {
          // Optional fallback if about:blank popup should not be shown
          if (ex instanceof BrowserAuthError && ex.errorCode === "popup_window_error") {
            accessToken = await getTokenWithDialogApi();
          } else {
            throw ex;
          }
        }
      }
    }

    // Need to show sign out button if Nested App Auth is not supported
    if (!Office.context.requirements.isSetSupported("NestedAppAuth")) {
      showSignOutButton();
    }

    output(accessToken);
  } catch (ex) {
    output(ex);
  }
}

function showSignOutButton() {
  document.getElementById("sign_out_button").style.display = "block";
}

async function getTokenWithDialogApi(isInternetExplorer?: boolean): Promise<string> {
  const accountContext = await getAccountContext();
  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      createLocalUrl(
        `${isInternetExplorer ? "dialogie.html" : "dialog.html"}?accountContext=${encodeURIComponent(JSON.stringify(accountContext))}`
      ),
      (result) => {
        result.value.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg: { message: string; origin: string | undefined }) => {
            const parsedMessage = JSON.parse(arg.message);
            resolve(parsedMessage.token);
            result.value.close();
          }
        );
      }
    );
  });
}

async function outputResultForDialogApi(isInternetExplorer?: boolean) {
  const token = await getTokenWithDialogApi(isInternetExplorer);
  output(token);
}
