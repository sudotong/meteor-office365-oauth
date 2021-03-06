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
  const params = {
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
let codeTimeouts = {};

const separateReferral = function (state) {
  let referral;
  try {
    const sep = encodeURIComponent('_ffr_');
    if (state && typeof state === 'string' && state.includes(sep)) {
      const [s, ref] = state.split(sep);
      state = s;
      referral = ref && decodeURIComponent(ref);
    }
  } catch (e) {
    console.log(`unable to sep referral`)
  }
  return [state, referral];
}

const getTokens = function (config, query) {

  let params = {
    scope: `offline_access openid profile User.Read Calendars.Read Calendars.ReadWrite`,
    code: query.code,
    client_id: config.clientId,
    client_secret: OAuth.openSecret(config.secret),
    redirect_uri: `${Meteor.absoluteUrl()}api/office365-auth`, // OAuth._redirectUri("office365", config).replace("?close", ""),
    state: query.state
  };

  // console.log({getTokens: 'getTokens', query});

  if (query.code && recentCodes[query.code]) {
    return recentCodes[query.code] == 'querying' ? null : recentCodes[query.code];
  }



  if (query.refresh_token) {
    params.grant_type = "refresh_token";
    params.refresh_token = query.refresh_token;
  } else {
    params.grant_type = "authorization_code";
    if (query.code) {
      recentCodes[query.code] = 'querying';
      codeTimeouts[query.code] = setTimeout(() => delete recentCodes[query.code], 11000);
    }
  }

  // TODO look at params.code_verifier for Azure AD login
  // if (query.code_verifier){
  // params.code_verifier = query.code_verifier;
  // }
  let initial_response = HTTP.post(`https://login.microsoftonline.com/${config.tenant || "common"}/oauth2/v2.0/token`, { headers: { Accept: "application/json", "User-Agent": userAgent }, params }).data;
  if (initial_response && query.code && !query.refresh_token) {
    if (codeTimeouts[query.code]) clearTimeout(codeTimeouts[query.code]);
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
    const identity = HTTP.get("https://graph.microsoft.com/v1.0/me", {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        "User-Agent": userAgent
      }
    }).data
    const emailKeys = ['mail', 'userPrincipalName', 'EmailAddress', 'mail'];
    if (identity) {
      emailKeys.forEach(key => {
        if (identity[key] && typeof identity[key] === 'string') identity[key] = identity[key].toLowerCase();
      })
    }
    return identity;
  } catch (error) {
    console.log('error getting office365 identity', error && error.message);
    return {};
  }
};

Meteor.methods({
  getEmail: function (refresh_token, access_token) {
    if (access_token) {
      let obj = getIdentity(access_token);
      return obj ? obj.mail || obj.userPrincipalName : null
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


  let referral;
  if (query) {
    const [newState, ref] = separateReferral(query.state);
    query.state = newState;
    referral = ref;
    // console.log(`office365_server:info`, `after state: ${query.state}, referral: ${referral}`);
  }


  try {
    data = getTokens(config, query);
  } catch (error) {
    console.log('Error getting tokens from office365 query', { message: error.message, config, query })
  }

  if (data) {
    const identity = getIdentity(data.access_token) || {};

    if (!identity.userPrincipalName) identity.userPrincipalName = identity.EmailAddress;

    const serviceData = {
      id: identity.id || identity.Id,
      accessToken: data.access_token,
      refreshToken: data.refresh_token,
      expiresAt: data.expires_in ? data.expires_in * 1000 + new Date().getTime() : null,
      scope: data.scope,
      displayName: identity.displayName || identity.DisplayName || identity.Alias,
      givenName: identity.givenName || identity.Alias || identity.displayName,
      surname: identity.surname,
      username: identity.userPrincipalName && identity.userPrincipalName.split("@")[0],
      userPrincipalName: identity.userPrincipalName,
      mail: identity.mail || identity.userPrincipalName,
      jobTitle: identity.jobTitle,
      mobilePhone: identity.mobilePhone,
      businessPhones: identity.businessPhones,
      officeLocation: identity.officeLocation,
      preferredLanguage: identity.preferredLanguage
    }
    const profile = { name: serviceData.givenName }
    if (referral) {
      serviceData.referral = referral;
    }
    return {
      serviceData,
      options: { profile }
    };
  } else {
    console.log('found no data for the user from query', { query });
    return {
      serviceData: {}
    }
  }

});

Office365.retrieveCredential = function (credentialToken, credentialSecret) {
  return OAuth.retrieveCredential(credentialToken, credentialSecret);
};