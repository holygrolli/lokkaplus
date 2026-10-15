import { AccessToken, TokenCredential, ClientSecretCredential, ClientCertificateCredential, DeviceCodeCredential, DeviceCodeInfo } from "@azure/identity";
import { PublicClientApplication, AccountInfo, ICachePlugin, ILoopbackClient, ServerAuthorizationCodeResponse } from "@azure/msal-node";
import { DataProtectionScope, PersistenceCachePlugin, PersistenceCreator } from "@azure/msal-node-extensions";
import { AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import { mkdirSync } from "fs";
import * as http from "http";
import { homedir } from "os";
import { dirname, join } from "path";
import jwt from "jsonwebtoken";
import { logger } from "./logger.js";
import { LokkaClientId, LokkaDefaultTenantId, LokkaDefaultRedirectUri } from "./constants.js";

// Constants
const ONE_HOUR_IN_MS = 60 * 60 * 1000; // One hour in milliseconds
const DEFAULT_TOKEN_CACHE_PATH = join(homedir(), ".lokkaplus", "msal-token-cache.json");
const GUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * Creates an MSAL cache plugin that persists tokens (incl. refresh token) across
 * server restarts, encrypted with the OS secret store (Windows DPAPI, macOS Keychain,
 * Linux libsecret). Returns undefined (in-memory cache only) if disabled or unavailable.
 */
async function createTokenCachePlugin(): Promise<ICachePlugin | undefined> {
  if (process.env.DISABLE_TOKEN_CACHE === "true") {
    logger.info("Persistent token cache disabled via DISABLE_TOKEN_CACHE");
    return undefined;
  }

  const cachePath = process.env.TOKEN_CACHE_PATH || DEFAULT_TOKEN_CACHE_PATH;
  try {
    mkdirSync(dirname(cachePath), { recursive: true, mode: 0o700 });
    const persistence = await PersistenceCreator.createPersistence({
      cachePath,
      dataProtectionScope: DataProtectionScope.CurrentUser,
      serviceName: "lokkaplus",
      accountName: "msal-token-cache",
      usePlaintextFileOnLinux: false,
    });
    logger.info(`Using persistent token cache at ${cachePath}`);
    return new PersistenceCachePlugin(persistence);
  } catch (error) {
    logger.error("Persistent token cache unavailable, falling back to in-memory cache", error);
    return undefined;
  }
}

// Helper function to parse JWT and extract scopes
function parseJwtScopes(token: string): string[] {
  try {
    // Decode JWT without verifying signature (we trust the token from Azure Identity)
    const decoded = jwt.decode(token) as any;
    
    if (!decoded || typeof decoded !== 'object') {
      logger.info("Failed to decode JWT token");
      return [];
    }

    // Extract scopes from the 'scp' claim (space-separated string)
    const scopesString = decoded.scp;
    if (typeof scopesString === 'string') {
      return scopesString.split(' ').filter(scope => scope.length > 0);
    }

    // Some tokens might have roles instead of scopes
    const roles = decoded.roles;
    if (Array.isArray(roles)) {
      return roles;
    }

    logger.info("No scopes found in JWT token");
    return [];
  } catch (error) {
    logger.error("Error parsing JWT token for scopes", error);
    return [];
  }
}

// Simple authentication provider that works with Azure Identity TokenCredential
export class TokenCredentialAuthProvider implements AuthenticationProvider {
  private credential: TokenCredential;

  constructor(credential: TokenCredential) {
    this.credential = credential;
  }

  async getAccessToken(): Promise<string> {
    const token = await this.credential.getToken("https://graph.microsoft.com/.default");
    if (!token) {
      throw new Error("Failed to acquire access token");
    }
    return token.token;
  }
}

export interface TokenBasedCredential extends TokenCredential {
  getToken(scopes: string | string[]): Promise<AccessToken | null>;
}

export class ClientProvidedTokenCredential implements TokenBasedCredential {
  private accessToken: string | undefined;
  private expiresOn: Date | undefined;
  constructor(accessToken?: string, expiresOn?: Date) {
    if (accessToken) {
      this.accessToken = accessToken;
      this.expiresOn = expiresOn || new Date(Date.now() + ONE_HOUR_IN_MS); // Default 1 hour
    } else {
      this.expiresOn = new Date(0); // Set to epoch to indicate no valid token
    }
  }
  async getToken(scopes: string | string[]): Promise<AccessToken | null> {
    if (!this.accessToken || !this.expiresOn || this.expiresOn <= new Date()) {
      logger.error("Access token is not available or has expired");
      return null;
    }

    return {
      token: this.accessToken,
      expiresOnTimestamp: this.expiresOn.getTime()
    };
  }  updateToken(accessToken: string, expiresOn?: Date): void {
    this.accessToken = accessToken;
    this.expiresOn = expiresOn || new Date(Date.now() + ONE_HOUR_IN_MS);
    logger.info("Access token updated successfully");
  }
  isExpired(): boolean {
    return !this.expiresOn || this.expiresOn <= new Date();
  }

  getExpirationTime(): Date {
    return this.expiresOn || new Date(0);
  }

  // Getter for access token (for internal use by AuthManager)
  getAccessToken(): string | undefined {
    return this.accessToken;
  }
}

/**
 * Custom loopback client that listens on "::" (IPv6 dual-stack).
 *
 * On macOS, /etc/hosts maps `localhost` to ::1 (IPv6) first.
 * @azure/msal-node's built-in LoopbackClient binds to 127.0.0.1 (IPv4 only),
 * so browsers that connect via ::1 get "connection refused" and the auth
 * callback is never delivered.  Binding to "::" accepts both IPv6 and IPv4
 * connections (dual-stack), fixing this transparently.
 */
class DualStackLoopbackClient implements ILoopbackClient {
  private server: http.Server | undefined;

  async listenForAuthCode(
    successTemplate?: string,
    errorTemplate?: string,
  ): Promise<ServerAuthorizationCodeResponse> {
    if (this.server) {
      throw new Error("Loopback server already exists");
    }

    this.server = http.createServer();

    // Wait for the server to actually start listening before returning the
    // outer promise so that getRedirectUri() is ready when MSAL polls it.
    await new Promise<void>((resolve, reject) => {
      this.server!.once("listening", resolve);
      this.server!.once("error", reject);
      this.server!.listen(0, "::"); // dual-stack: handles both ::1 (IPv6) and 127.0.0.1 (IPv4)
    });

    return new Promise<ServerAuthorizationCodeResponse>((resolve, reject) => {
      this.server!.on("request", (req: http.IncomingMessage, res: http.ServerResponse) => {
        const url = req.url;
        if (!url) {
          res.end(errorTemplate ?? "Error occurred loading redirect URL");
          reject(new Error("Unable to load redirect URL"));
          return;
        }
        if (url === "/") {
          // Browser followed the 302 redirect after code delivery — show success page.
          res.end(successTemplate ?? "Auth code was successfully acquired. You can close this window now.");
          return;
        }
        const redirectUri = this.getRedirectUri();
        const parsedUrl = new URL(url, redirectUri);
        const params: Record<string, string> = {};
        parsedUrl.searchParams.forEach((v, k) => { params[k] = v; });

        if (params["code"]) {
          // Redirect browser to "/" so the success page is shown.
          res.writeHead(302, { location: redirectUri });
          res.end();
        }
        if (params["error"]) {
          res.end(errorTemplate ?? `Error: ${params["error"]}`);
        }
        resolve(params as ServerAuthorizationCodeResponse);
      });
    });
  }

  getRedirectUri(): string {
    if (!this.server?.listening) {
      // Throw a plain error; MSAL's waitForRedirectUri will reject if it's not
      // the expected NodeAuthError, but by the time MSAL first polls (100 ms)
      // the server will already be listening since we awaited the listen event.
      throw new Error("Loopback server is not listening yet");
    }
    const address = this.server.address() as { port: number };
    return `http://localhost:${address.port}`;
  }

  closeServer(): void {
    if (this.server) {
      this.server.close();
      try { (this.server as any).closeAllConnections?.(); } catch { /* ignore */ }
      this.server.unref();
      this.server = undefined;
    }
  }
}

/**
 * Interactive browser credential backed directly by @azure/msal-node.
 *
 * Fixes two issues present in @azure/identity's InteractiveBrowserCredential:
 *   1. The built-in openBrowser uses `open(url, { wait: true })` which blocks
 *      until the browser window is closed, so token exchange never happens.
 *   2. The built-in LoopbackClient listens on 127.0.0.1 (IPv4) only, which
 *      fails on systems where `localhost` resolves to ::1 (IPv6) first.
 */
class MsalInteractiveCredential implements TokenCredential {
  private pcaPromise: Promise<PublicClientApplication> | null = null;
  private account: AccountInfo | null = null;
  private tenantId: string;
  private clientId: string;

  constructor(tenantId: string, clientId: string) {
    this.tenantId = tenantId;
    this.clientId = clientId;
  }

  private getPca(): Promise<PublicClientApplication> {
    if (!this.pcaPromise) {
      this.pcaPromise = createTokenCachePlugin().then(cachePlugin => new PublicClientApplication({
        auth: {
          clientId: this.clientId,
          authority: `https://login.microsoftonline.com/${this.tenantId}`,
        },
        cache: cachePlugin ? { cachePlugin } : undefined,
        system: {
          loggerOptions: {
            loggerCallback: () => { /* suppress MSAL verbose logs */ },
            piiLoggingEnabled: false,
            logLevel: 3, // Error
          },
        },
      }));
    }
    return this.pcaPromise;
  }

  // Restores the signed-in account from the persistent cache after a server restart.
  private async findCachedAccount(pca: PublicClientApplication): Promise<AccountInfo | null> {
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (GUID_PATTERN.test(this.tenantId)) {
      return accounts.find(a => a.tenantId?.toLowerCase() === this.tenantId.toLowerCase()) ?? null;
    }
    return accounts[0] ?? null;
  }

  async getToken(scopes: string | string[]): Promise<AccessToken | null> {
    const scopeArray = Array.isArray(scopes) ? scopes : [scopes];
    const pca = await this.getPca();

    if (!this.account) {
      this.account = await this.findCachedAccount(pca);
      if (this.account) {
        logger.info(`Restored cached account ${this.account.username} from token cache`);
      }
    }

    // Try silent authentication with a cached account first (uses refresh token if needed).
    if (this.account) {
      try {
        const silentResult = await pca.acquireTokenSilent({
          scopes: scopeArray,
          account: this.account,
          forceRefresh: false,
        });
        if (silentResult?.accessToken) {
          return {
            token: silentResult.accessToken,
            expiresOnTimestamp: silentResult.expiresOn?.getTime() ?? Date.now() + 3600_000,
          };
        }
      } catch (error) {
        logger.info("Silent token acquisition failed, falling back to interactive login", error instanceof Error ? error.message : error);
      }
    }

    // Interactive auth: open browser (without wait:true) + dual-stack loopback.
    const result = await pca.acquireTokenInteractive({
      scopes: scopeArray,
      loopbackClient: new DualStackLoopbackClient(),
      openBrowser: async (url: string) => {
        // Dynamic import of `open` (ESM-only package).
        const { default: open } = await import("open");
        // Do NOT use { wait: true } — that blocks until the browser process
        // exits, delaying token exchange until the user closes the browser.
        await open(url, { wait: false, newInstance: false });
      },
    });

    if (!result?.accessToken) {
      throw new Error("Interactive authentication did not return an access token");
    }

    this.account = result.account;
    return {
      token: result.accessToken,
      expiresOnTimestamp: result.expiresOn?.getTime() ?? Date.now() + 3600_000,
    };
  }
}

export enum AuthMode {
  ClientCredentials = "client_credentials",
  ClientProvidedToken = "client_provided_token", 
  Interactive = "interactive",
  Certificate = "certificate"
}

export interface AuthConfig {
  mode: AuthMode;
  tenantId?: string;
  clientId?: string;
  clientSecret?: string;
  accessToken?: string;
  expiresOn?: Date;
  redirectUri?: string;
  certificatePath?: string;
  certificatePassword?: string;
}

export class AuthManager {
  private credential: TokenCredential | null = null;
  private config: AuthConfig;

  constructor(config: AuthConfig) {
    this.config = config;
  }

  async initialize(): Promise<void> {
    switch (this.config.mode) {
      case AuthMode.ClientCredentials:
        if (!this.config.tenantId || !this.config.clientId || !this.config.clientSecret) {
          throw new Error("Client credentials mode requires tenantId, clientId, and clientSecret");
        }
        logger.info("Initializing Client Credentials authentication");
        this.credential = new ClientSecretCredential(
          this.config.tenantId,
          this.config.clientId,
          this.config.clientSecret
        );
        break;

      case AuthMode.ClientProvidedToken:
        logger.info("Initializing Client Provided Token authentication");
        this.credential = new ClientProvidedTokenCredential(
          this.config.accessToken,
          this.config.expiresOn
        );
        break;
        
      case AuthMode.Certificate:
        if (!this.config.tenantId || !this.config.clientId || !this.config.certificatePath) {
          throw new Error("Certificate mode requires tenantId, clientId, and certificatePath");
        }
        logger.info("Initializing Certificate authentication");
        this.credential = new ClientCertificateCredential(this.config.tenantId, this.config.clientId, {
          certificatePath: this.config.certificatePath,
          certificatePassword: this.config.certificatePassword
        });
        break;

      case AuthMode.Interactive:
        // Use defaults if not provided
        const tenantId = this.config.tenantId || LokkaDefaultTenantId;
        const clientId = this.config.clientId || LokkaClientId;
        
        logger.info(`Initializing Interactive authentication with tenant ID: ${tenantId}, client ID: ${clientId}`);
        
        // Use MsalInteractiveCredential instead of @azure/identity's
        // InteractiveBrowserCredential to avoid two known issues:
        //   1. InteractiveBrowserCredential opens the browser with { wait: true },
        //      blocking token exchange until the window is closed.
        //   2. Its loopback server binds to 127.0.0.1 only, which fails when
        //      the browser resolves `localhost` to ::1 (IPv6) first on macOS.
        this.credential = new MsalInteractiveCredential(tenantId, clientId);
        break;

      default:
        throw new Error(`Unsupported authentication mode: ${this.config.mode}`);
    }

    // Test the credential (skip interactive — auth is deferred to first tool call)
    if (this.config.mode !== AuthMode.Interactive) {
      await this.testCredential();
    }
  }

  updateAccessToken(accessToken: string, expiresOn?: Date): void {
    if (this.config.mode === AuthMode.ClientProvidedToken && this.credential instanceof ClientProvidedTokenCredential) {
      this.credential.updateToken(accessToken, expiresOn);
    } else {
      throw new Error("Token update only supported in client provided token mode");
    }
  }
  private async testCredential(): Promise<void> {
    if (!this.credential) {
      throw new Error("Credential not initialized");
    }

    // Skip testing if ClientProvidedToken mode has no initial token
    if (this.config.mode === AuthMode.ClientProvidedToken && !this.config.accessToken) {
      logger.info("Skipping initial credential test as no token was provided at startup.");
      return;
    }

    try {
      const token = await this.credential.getToken("https://graph.microsoft.com/.default");
      if (!token) {
        throw new Error("Failed to acquire token");
      }
      logger.info("Authentication successful");
    } catch (error) {
      logger.error("Authentication test failed", error);
      throw error;
    }
  }
  getGraphAuthProvider(): TokenCredentialAuthProvider {
    if (!this.credential) {
      throw new Error("Authentication not initialized");
    }

    return new TokenCredentialAuthProvider(this.credential);
  }

  getAzureCredential(): TokenCredential {
    if (!this.credential) {
      throw new Error("Authentication not initialized");
    }
    return this.credential;
  }

  getAuthMode(): AuthMode {
    return this.config.mode;
  }

  isClientCredentials(): boolean {
    return this.config.mode === AuthMode.ClientCredentials;
  }

  isClientProvidedToken(): boolean {
    return this.config.mode === AuthMode.ClientProvidedToken;
  }

  isInteractive(): boolean {
    return this.config.mode === AuthMode.Interactive;
  }

  async getTokenStatus(): Promise<{ isExpired: boolean; expiresOn?: Date; scopes?: string[] }> {
    if (this.credential instanceof ClientProvidedTokenCredential) {
      const tokenStatus = {
        isExpired: this.credential.isExpired(),
        expiresOn: this.credential.getExpirationTime()
      };

      // If we have a valid token, parse it to extract scopes
      if (!tokenStatus.isExpired) {
        const accessToken = this.credential.getAccessToken();
        if (accessToken) {
          try {
            const scopes = parseJwtScopes(accessToken);
            return {
              ...tokenStatus,
              scopes: scopes
            };
          } catch (error) {
            logger.error("Error parsing token scopes in getTokenStatus", error);
            return tokenStatus;
          }
        }
      }

      return tokenStatus;
    } else if (this.credential) {
      // For other credential types, try to get a fresh token and parse it
      try {
        const accessToken = await this.credential.getToken("https://graph.microsoft.com/.default");
        if (accessToken && accessToken.token) {
          const scopes = parseJwtScopes(accessToken.token);
          return {
            isExpired: false,
            expiresOn: new Date(accessToken.expiresOnTimestamp),
            scopes: scopes
          };
        }
      } catch (error) {
        logger.error("Error getting token for scope parsing", error);
      }
    }
    
    return { isExpired: false };
  }
}
