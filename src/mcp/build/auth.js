import { ClientSecretCredential, ClientCertificateCredential } from "@azure/identity";
import { PublicClientApplication } from "@azure/msal-node";
import * as http from "http";
import jwt from "jsonwebtoken";
import { logger } from "./logger.js";
import { LokkaClientId, LokkaDefaultTenantId } from "./constants.js";
// Constants
const ONE_HOUR_IN_MS = 60 * 60 * 1000; // One hour in milliseconds
// Helper function to parse JWT and extract scopes
function parseJwtScopes(token) {
    try {
        // Decode JWT without verifying signature (we trust the token from Azure Identity)
        const decoded = jwt.decode(token);
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
    }
    catch (error) {
        logger.error("Error parsing JWT token for scopes", error);
        return [];
    }
}
// Simple authentication provider that works with Azure Identity TokenCredential
export class TokenCredentialAuthProvider {
    credential;
    constructor(credential) {
        this.credential = credential;
    }
    async getAccessToken() {
        const token = await this.credential.getToken("https://graph.microsoft.com/.default");
        if (!token) {
            throw new Error("Failed to acquire access token");
        }
        return token.token;
    }
}
export class ClientProvidedTokenCredential {
    accessToken;
    expiresOn;
    constructor(accessToken, expiresOn) {
        if (accessToken) {
            this.accessToken = accessToken;
            this.expiresOn = expiresOn || new Date(Date.now() + ONE_HOUR_IN_MS); // Default 1 hour
        }
        else {
            this.expiresOn = new Date(0); // Set to epoch to indicate no valid token
        }
    }
    async getToken(scopes) {
        if (!this.accessToken || !this.expiresOn || this.expiresOn <= new Date()) {
            logger.error("Access token is not available or has expired");
            return null;
        }
        return {
            token: this.accessToken,
            expiresOnTimestamp: this.expiresOn.getTime()
        };
    }
    updateToken(accessToken, expiresOn) {
        this.accessToken = accessToken;
        this.expiresOn = expiresOn || new Date(Date.now() + ONE_HOUR_IN_MS);
        logger.info("Access token updated successfully");
    }
    isExpired() {
        return !this.expiresOn || this.expiresOn <= new Date();
    }
    getExpirationTime() {
        return this.expiresOn || new Date(0);
    }
    // Getter for access token (for internal use by AuthManager)
    getAccessToken() {
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
class DualStackLoopbackClient {
    server;
    async listenForAuthCode(successTemplate, errorTemplate) {
        if (this.server) {
            throw new Error("Loopback server already exists");
        }
        this.server = http.createServer();
        // Wait for the server to actually start listening before returning the
        // outer promise so that getRedirectUri() is ready when MSAL polls it.
        await new Promise((resolve, reject) => {
            this.server.once("listening", resolve);
            this.server.once("error", reject);
            this.server.listen(0, "::"); // dual-stack: handles both ::1 (IPv6) and 127.0.0.1 (IPv4)
        });
        return new Promise((resolve, reject) => {
            this.server.on("request", (req, res) => {
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
                const params = {};
                parsedUrl.searchParams.forEach((v, k) => { params[k] = v; });
                if (params["code"]) {
                    // Redirect browser to "/" so the success page is shown.
                    res.writeHead(302, { location: redirectUri });
                    res.end();
                }
                if (params["error"]) {
                    res.end(errorTemplate ?? `Error: ${params["error"]}`);
                }
                resolve(params);
            });
        });
    }
    getRedirectUri() {
        if (!this.server?.listening) {
            // Throw a plain error; MSAL's waitForRedirectUri will reject if it's not
            // the expected NodeAuthError, but by the time MSAL first polls (100 ms)
            // the server will already be listening since we awaited the listen event.
            throw new Error("Loopback server is not listening yet");
        }
        const address = this.server.address();
        return `http://localhost:${address.port}`;
    }
    closeServer() {
        if (this.server) {
            this.server.close();
            try {
                this.server.closeAllConnections?.();
            }
            catch { /* ignore */ }
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
class MsalInteractiveCredential {
    pca;
    account = null;
    tenantId;
    constructor(tenantId, clientId) {
        this.tenantId = tenantId;
        this.pca = new PublicClientApplication({
            auth: {
                clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
            },
            system: {
                loggerOptions: {
                    loggerCallback: () => { },
                    piiLoggingEnabled: false,
                    logLevel: 3, // Error
                },
            },
        });
    }
    async getToken(scopes) {
        const scopeArray = Array.isArray(scopes) ? scopes : [scopes];
        // Try silent authentication with a cached account first.
        if (this.account) {
            try {
                const silentResult = await this.pca.acquireTokenSilent({
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
            }
            catch {
                // Silent auth failed — fall through to interactive.
            }
        }
        // Interactive auth: open browser (without wait:true) + dual-stack loopback.
        const result = await this.pca.acquireTokenInteractive({
            scopes: scopeArray,
            loopbackClient: new DualStackLoopbackClient(),
            openBrowser: async (url) => {
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
export var AuthMode;
(function (AuthMode) {
    AuthMode["ClientCredentials"] = "client_credentials";
    AuthMode["ClientProvidedToken"] = "client_provided_token";
    AuthMode["Interactive"] = "interactive";
    AuthMode["Certificate"] = "certificate";
})(AuthMode || (AuthMode = {}));
export class AuthManager {
    credential = null;
    config;
    constructor(config) {
        this.config = config;
    }
    async initialize() {
        switch (this.config.mode) {
            case AuthMode.ClientCredentials:
                if (!this.config.tenantId || !this.config.clientId || !this.config.clientSecret) {
                    throw new Error("Client credentials mode requires tenantId, clientId, and clientSecret");
                }
                logger.info("Initializing Client Credentials authentication");
                this.credential = new ClientSecretCredential(this.config.tenantId, this.config.clientId, this.config.clientSecret);
                break;
            case AuthMode.ClientProvidedToken:
                logger.info("Initializing Client Provided Token authentication");
                this.credential = new ClientProvidedTokenCredential(this.config.accessToken, this.config.expiresOn);
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
    updateAccessToken(accessToken, expiresOn) {
        if (this.config.mode === AuthMode.ClientProvidedToken && this.credential instanceof ClientProvidedTokenCredential) {
            this.credential.updateToken(accessToken, expiresOn);
        }
        else {
            throw new Error("Token update only supported in client provided token mode");
        }
    }
    async testCredential() {
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
        }
        catch (error) {
            logger.error("Authentication test failed", error);
            throw error;
        }
    }
    getGraphAuthProvider() {
        if (!this.credential) {
            throw new Error("Authentication not initialized");
        }
        return new TokenCredentialAuthProvider(this.credential);
    }
    getAzureCredential() {
        if (!this.credential) {
            throw new Error("Authentication not initialized");
        }
        return this.credential;
    }
    getAuthMode() {
        return this.config.mode;
    }
    isClientCredentials() {
        return this.config.mode === AuthMode.ClientCredentials;
    }
    isClientProvidedToken() {
        return this.config.mode === AuthMode.ClientProvidedToken;
    }
    isInteractive() {
        return this.config.mode === AuthMode.Interactive;
    }
    async getTokenStatus() {
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
                    }
                    catch (error) {
                        logger.error("Error parsing token scopes in getTokenStatus", error);
                        return tokenStatus;
                    }
                }
            }
            return tokenStatus;
        }
        else if (this.credential) {
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
            }
            catch (error) {
                logger.error("Error getting token for scope parsing", error);
            }
        }
        return { isExpired: false };
    }
}
