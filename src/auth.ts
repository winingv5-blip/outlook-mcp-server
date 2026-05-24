/**
 * Microsoft Graph authentication via MSAL Node (Device Code Flow).
 *
 * On first run the server prints a short URL + one-time code to stderr.
 * The user visits the URL in a browser, enters the code, and signs in.
 * The resulting token is cached to TOKEN_CACHE_FILE and refreshed silently
 * on subsequent runs.
 *
 * Required environment variables:
 *   OUTLOOK_CLIENT_ID   – Azure app registration client (application) ID
 *   OUTLOOK_TENANT_ID   – Azure tenant ID, or "common" / "consumers"
 */

import {
  PublicClientApplication,
  type Configuration,
  type AuthenticationResult,
  type DeviceCodeRequest,
  type SilentFlowRequest,
  type AccountInfo,
} from "@azure/msal-node";
import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";
import { MAIL_SCOPES, TOKEN_CACHE_FILE } from "./constants.js";

// ─── Token cache (file-backed) ────────────────────────────────────────────────

const CACHE_PATH = path.join(os.homedir(), TOKEN_CACHE_FILE);

function readCacheFile(): string {
  try {
    return fs.existsSync(CACHE_PATH) ? fs.readFileSync(CACHE_PATH, "utf-8") : "";
  } catch {
    return "";
  }
}

function writeCacheFile(data: string): void {
  try {
    fs.writeFileSync(CACHE_PATH, data, { mode: 0o600 });
  } catch (err) {
    console.error("[auth] Failed to write token cache:", err);
  }
}

// ─── MSAL client ──────────────────────────────────────────────────────────────

let _pca: PublicClientApplication | null = null;

function getPca(): PublicClientApplication {
  if (_pca) return _pca;

  const clientId = process.env.OUTLOOK_CLIENT_ID;
  const tenantId = process.env.OUTLOOK_TENANT_ID ?? "common";

  if (!clientId) {
    throw new Error(
      "OUTLOOK_CLIENT_ID environment variable is required.\n" +
        "Register an app at https://portal.azure.com and set the variable."
    );
  }

  const serializedCache = readCacheFile();

  const config: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: {
        beforeCacheAccess: async (ctx) => {
          ctx.tokenCache.deserialize(serializedCache);
        },
        afterCacheAccess: async (ctx) => {
          if (ctx.cacheHasChanged) {
            writeCacheFile(ctx.tokenCache.serialize());
          }
        },
      },
    },
  };

  _pca = new PublicClientApplication(config);
  return _pca;
}

// ─── Token acquisition ────────────────────────────────────────────────────────

async function getSilentToken(account: AccountInfo): Promise<string | null> {
  const pca = getPca();
  const request: SilentFlowRequest = {
    account,
    scopes: MAIL_SCOPES,
  };
  try {
    const result: AuthenticationResult = await pca.acquireTokenSilent(request);
    return result.accessToken;
  } catch {
    return null;
  }
}

async function getDeviceCodeToken(): Promise<string> {
  const pca = getPca();
  const request: DeviceCodeRequest = {
    scopes: MAIL_SCOPES,
    deviceCodeCallback: (response) => {
      // Print to stderr so it doesn't pollute the MCP stdio channel
      console.error("\n════════════════════════════════════════");
      console.error("  Outlook MCP — Sign in required");
      console.error("════════════════════════════════════════");
      console.error(`  1. Open:  ${response.verificationUri}`);
      console.error(`  2. Enter: ${response.userCode}`);
      console.error("════════════════════════════════════════\n");
    },
  };
  const result: AuthenticationResult = await pca.acquireTokenByDeviceCode(request);
  return result.accessToken;
}

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Returns a valid Microsoft Graph access token, refreshing silently if possible
 * or triggering device-code flow for first-time auth.
 */
export async function getAccessToken(): Promise<string> {
  const pca = getPca();
  const accounts = await pca.getTokenCache().getAllAccounts();

  if (accounts.length > 0) {
    const token = await getSilentToken(accounts[0]);
    if (token) return token;
  }

  // Silent failed or no cached account → device-code flow
  return getDeviceCodeToken();
}

/**
 * Clears the cached token (forces re-authentication on next call).
 */
export function clearTokenCache(): void {
  try {
    if (fs.existsSync(CACHE_PATH)) {
      fs.unlinkSync(CACHE_PATH);
    }
    _pca = null;
  } catch (err) {
    console.error("[auth] Failed to clear token cache:", err);
  }
}
