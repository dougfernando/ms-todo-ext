
import { getPreferenceValues, OAuth } from "@raycast/api";
import fetch from "node-fetch";

interface Preferences {
  clientId: string;
}

const scope = "Tasks.ReadWrite offline_access";

export const oauthClient = new OAuth.PKCEClient({
  redirectMethod: OAuth.RedirectMethod.Web,
  providerName: "Microsoft",
  providerId: "microsoft",
  providerIcon: "microsoft.png",
  description: "Connect your Microsoft account to get started.",
});

export async function authorize(): Promise<void> {
  const { clientId } = getPreferenceValues<Preferences>();
  const tokenSet = await oauthClient.getTokens();
  if (tokenSet?.accessToken) {
    if (tokenSet.refreshToken && tokenSet.isExpired()) {
      const newTokens = await refreshTokens(tokenSet.refreshToken);
      await oauthClient.setTokens(newTokens);
    }
    return;
  }

  const authRequest = await oauthClient.authorizationRequest({
    endpoint: "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
    clientId,
    scope,
  });
  console.log("Auth Request before authorize:", JSON.stringify(authRequest));
  const { authorizationCode } = await oauthClient.authorize(authRequest);
  console.log("Auth Request after authorize:", JSON.stringify(authRequest));
  const newTokens = await fetchTokens(authRequest, authorizationCode);
  await oauthClient.setTokens(newTokens);
}

async function fetchTokens(authRequest: OAuth.AuthorizationRequest, authCode: string): Promise<OAuth.TokenResponse> {
  console.log("Auth Request in fetchTokens:", JSON.stringify(authRequest));
  const { clientId } = getPreferenceValues<Preferences>();
  const response = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      code: authCode,
      redirect_uri: authRequest.redirectURI,
      code_verifier: authRequest.codeVerifier,
      grant_type: "authorization_code",
    }),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(error);
  }

  return (await response.json()) as OAuth.TokenResponse;
}

async function refreshTokens(refreshToken: string): Promise<OAuth.TokenResponse> {
  const { clientId } = getPreferenceValues<Preferences>();
  const response = await fetch("https://login.microsoftonline.com/common/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      refresh_token: refreshToken,
      grant_type: "refresh_token",
      scope,
    }),
  });

  if (!response.ok) {
    await oauthClient.removeTokens();
    throw new Error("Failed to refresh tokens");
  }

  const tokenResponse = (await response.json()) as OAuth.TokenResponse;
  tokenResponse.refresh_token = tokenResponse.refresh_token ?? refreshToken;
  return tokenResponse;
}

export async function getAccessToken(): Promise<string> {
  const tokenSet = await oauthClient.getTokens();
  if (!tokenSet?.accessToken) {
    throw new Error("Not authenticated");
  }
  if (tokenSet.refreshToken && tokenSet.isExpired()) {
    const newTokens = await refreshTokens(tokenSet.refreshToken);
    await oauthClient.setTokens(newTokens);
    return newTokens.access_token;
  }
  return tokenSet.accessToken;
}
