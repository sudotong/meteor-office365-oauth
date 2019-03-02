/*globals Office365, OAuth */
"use strict";

Office365 = {};

let userAgent = "Meteor";
if (Meteor.release) {
  userAgent += `/${Meteor.release}`;
}

const getAccessFromRefresh = function (refreshToken) {
  let response;
  const config = ServiceConfiguration.configurations.findOne({ service: "office365" });
  if (!config) throw new ServiceConfiguration.ConfigError();
  let params = {
    scope: `offline_access openid profile User.Read Calendars.Read Calendars.ReadWrite`,
    client_id: config.clientId,
    client_secret: OAuth.openSecret(config.secret),
    redirect_uri: OAuth._redirectUri("office365", config).replace("?close", "")
  };
  params.grant_type = "refresh_token";
  params.refresh_token = refreshToken;

  response = HTTP.post(`https://login.microsoftonline.com/${config.tenant || "common"}/oauth2/v2.0/token`, {
    headers: { Accept: "application/json", "User-Agent": userAgent }, params
  }).data;

  return response.access_token;
}

let recentCodes = {};

const getTokens = function (config, query) {

  let params = {
    scope: `offline_access openid profile User.Read Calendars.Read Calendars.ReadWrite`,
    code: query.code, 
    client_id: config.clientId,
    client_secret: OAuth.openSecret(config.secret),
    redirect_uri: `${Meteor.absoluteUrl()}api/office365-auth`, // OAuth._redirectUri("office365", config).replace("?close", ""),
    state: query.state
  };
  if (query.code && recentCodes[query.code]) return recentCodes[query.code];

  

  if (query.refresh_token) {
    params.grant_type = "refresh_token";
    params.refresh_token = query.refresh_token;
  } else {
    params.grant_type = "authorization_code";
  }

  let initial_response = HTTP.post(`https://login.microsoftonline.com/${config.tenant || "common"}/oauth2/v2.0/token`, { headers: { Accept: "application/json", "User-Agent": userAgent }, params }).data;
  if (initial_response && query.code && !query.refresh_token){
    recentCodes[query.code] = initial_response;
    setTimeout(() => delete recentCodes[query.code], 5000);
  }
  // console.log({initial_response});

  if (initial_response && params.grant_type === "authorization_code" && initial_response.refresh_token) {
    // Meteor.users.update({'services.office365.refreshToken': }, { $set: { 'services.office365.refreshToken': initial_response.refresh_token } })
  }

  if (initial_response && params.grant_type === "refresh_token") {
    // Meteor.users.update({'services.office365.refreshToken': }, { $set: { 'services.office365.refreshToken': false } })
  }

  return initial_response;
};

const getIdentity = function (accessToken) {
  try {
    if (!accessToken) return {};
    return HTTP.get("https://graph.microsoft.com/v1.0/me", {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        "User-Agent": userAgent
      }
    }).data;
  } catch (error) {
    console.log(error.message)
    return {};
  }
};

Meteor.methods({
  getEmail: function (refresh_token, access_token) {
    console.log("this is being called")
    if (access_token) {
      let obj = getIdentity(access_token);
      return obj.mail || obj.userPrincipalName
    } else {
      let obj = getIdentity(getAccessFromRefresh(refresh_token));
      return obj.mail || obj.userPrincipalName
    }
  }
})

OAuth.registerService("office365", 2, null, function (query, other) {
  let data;

  /**
   * Make sure we have a config object for subsequent use (boilerplate)
   */
  const config = ServiceConfiguration.configurations.findOne({ service: "office365" });
  if (!config) throw new ServiceConfiguration.ConfigError();

  try {
    data = getTokens(config, query);
  } catch (error) {
    console.log('Error getting tokens from office365 query', error.message)
  }

  if (data) {
    const identity = getIdentity(data.access_token) || {};

    if (!identity.userPrincipalName) identity.userPrincipalName = identity.EmailAddress;

    const serviceData = {
      id: identity.id || identity.Id,
      accessToken: data.access_token,
      refreshToken: data.refresh_token,
      expiresAt: data.expires_in
        ? data.expires_in * 1000 + new Date().getTime()
        : null,
      scope: data.scope,
      displayName: identity.displayName ||
        identity.DisplayName ||
        identity.Alias,
      givenName: identity.givenName || identity.Alias || identity.displayName,
      surname: identity.surname,
      username: identity.userPrincipalName &&
        identity.userPrincipalName.split("@")[0],
      userPrincipalName: identity.userPrincipalName,
      mail: identity.mail || identity.userPrincipalName,
      jobTitle: identity.jobTitle,
      mobilePhone: identity.mobilePhone,
      businessPhones: identity.businessPhones,
      officeLocation: identity.officeLocation,
      preferredLanguage: identity.preferredLanguage
    }
    return {
      serviceData,
      options: { profile: { name: serviceData.givenName } }
    };
  } else {
    return {
      serviceData: {}
    }
  }

});

Office365.retrieveCredential = function (credentialToken, credentialSecret) {
  return OAuth.retrieveCredential(credentialToken, credentialSecret);
};