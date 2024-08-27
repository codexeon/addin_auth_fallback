/* global console, Office */

import { getTokenRequest, AccountContext, ensurePublicClient } from "./msalcommon";
import { createLocalUrl } from "./util";

// read querystring parameter
function getQueryParameter(param: string) {
  const params = new URLSearchParams(window.location.search);
  return params.get(param);
}

export async function initializeMsal() {
  const publicClientApp = await ensurePublicClient();
  try {
    if (getQueryParameter("logout") === "1") {
      await publicClientApp.logoutRedirect();
      return;
    }
    const result = await publicClientApp.handleRedirectPromise();

    if (result) {
      publicClientApp.setActiveAccount(result.account);
      await Office.onReady();
      Office.context.ui.messageParent(JSON.stringify({ token: result.accessToken }));
      return;
    }
  } catch (ex) {
    await Office.onReady();
    Office.context.ui.messageParent(JSON.stringify({ error: ex.name }));
    return;
  }

  const accountContextString = getQueryParameter("accountContext");
  let accountContext: AccountContext;
  if (accountContextString) {
    accountContext = JSON.parse(accountContextString);
  }
  const request = await getTokenRequest(accountContext);
  publicClientApp.loginRedirect({
    ...request,
    redirectUri: createLocalUrl("dialog.html"),
  });
}

initializeMsal();
