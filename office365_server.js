/*globals Office365, OAuth */
'use strict';

Office365 = {};

let userAgent = 'Meteor';
if (Meteor.release) { userAgent += `/${ Meteor.release }`; }

const getTokens = function(query) {
  const config = ServiceConfiguration.configurations.findOne({service: 'office365'});
  if (!config) { throw new ServiceConfiguration.ConfigError(); }

  const redirectUri = OAuth._redirectUri('office365', config).replace('?close', '');

  // https://github.com/microsoftgraph/microsoft-graph-docs/blob/master/concepts/auth_v2_user.md

  let response;
  try {
    response = HTTP.post(
      `https://login.microsoftonline.com/${ config.tenant || 'common' }/oauth2/v2.0/token`, {
        headers: {
          Accept: 'application/json',
          'User-Agent': userAgent
        },
        params: {
          grant_type: 'authorization_code',
          code: query.code,
          client_id: config.clientId,
          client_secret: OAuth.openSecret(config.secret),
          redirect_uri: redirectUri,
          state: query.state
        }
      });
  } catch (error) {
    throw _.extend(new Error(`Failed to complete OAuth handshake with Microsoft Office365. ${ error.message }`), {response: error.response});
  }
  if (response.data.error) {
    throw new Error(`Failed to complete OAuth handshake with Microsoft Office365. ${ response.data.error }`);
  } else {
    return response.data;
  }
};

const getIdentity = function(accessToken) {
  try {
    return HTTP.get(
      'https://outlook.office.com/api/v2.0/me', {
        headers: {
          Authorization: `Bearer ${ accessToken }`,
          Accept: 'application/json',
          'User-Agent': userAgent
        }
      }).data;
  } catch (error) {
    throw _.extend(new Error(`Failed to fetch identity from Microsoft Office365. ${ error.message }`), {response: error.response});
  }
};

OAuth.registerService('office365', 2, null, function(query) {
  const data = getTokens(query);
  const identity = getIdentity(data.access_token);
  return {
    serviceData: {
      id: identity.id || identity.Id,
      accessToken: OAuth.sealSecret(data.access_token),
      refreshToken: data.refresh_token ? OAuth.sealSecret(data.refresh_token) : null,
      expiresAt: data.expires_in ? data.expires_in*1000 + new Date().getTime() : null,
      scope: data.scope,
      displayName: identity.displayName || identity.DisplayName,
      givenName: identity.givenName,
      surname: identity.surname,
      username: identity.userPrincipalName && identity.userPrincipalName.split('@')[0],
      userPrincipalName: identity.userPrincipalName,
      mail: identity.mail,
      jobTitle: identity.jobTitle,
      mobilePhone: identity.mobilePhone,
      businessPhones: identity.businessPhones,
      officeLocation: identity.officeLocation,
      preferredLanguage: identity.preferredLanguage
    },
    options: {profile: {name: identity.givenName}}
  };
});

Office365.retrieveCredential = function(credentialToken, credentialSecret) {
  return OAuth.retrieveCredential(credentialToken, credentialSecret);
};
