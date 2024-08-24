/* global console */

import {
  AccountInfo,
  Configuration,
  createNestablePublicClientApplication,
  createStandardPublicClientApplication,
  LogLevel,
  PublicClientApplication,
  type RedirectRequest,
} from "@azure/msal-browser";

export function createLocalUrl(path: string) {
  return `${window.location.origin}/${path}`;
}
const clientId = "148b0448-c6ab-4d8e-adb2-a0f2696966d2";
export const msalConfig: Configuration = {
  auth: {
    clientId,
    redirectUri: createLocalUrl("auth.html"),
    postLogoutRedirectUri: createLocalUrl("auth.html"),
  },
  cache: {
    cacheLocation: "localStorage",
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) {
          return;
        }
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
    },
  },
};

export async function getTokenRequest(accountContext?: AccountContext): Promise<RedirectRequest> {
  const account = await getAccountFromContext(accountContext);
  let additionalProperties: Partial<RedirectRequest> = {};
  if (account) {
    additionalProperties = { account };
  } else if (accountContext) {
    additionalProperties = {
      loginHint: accountContext.loginHint,
    };
  } else {
    additionalProperties = { prompt: "select_account" };
  }
  return { scopes: ["user.read"], ...additionalProperties };
}
let _publicClientApp: PublicClientApplication;
export async function ensurePublicClient() {
  if (!_publicClientApp) {
    _publicClientApp = await createNestablePublicClientApplication(msalConfig);
  }
  return _publicClientApp;
}

export type AccountContext = {
  loginHint?: string;
  tenantId?: string;
  localAccountId?: string;
};

export async function getAccountFromContext(accountContext?: AccountContext): Promise<AccountInfo | null> {
  const pca = await ensurePublicClient();
  if (!accountContext) {
    return pca.getActiveAccount();
  }

  return pca.getAccount({
    username: accountContext.loginHint,
    tenantId: accountContext.tenantId,
    localAccountId: accountContext.localAccountId,
  });
}
