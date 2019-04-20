/*globals Office365, OAuth */
import { Tracker } from 'meteor/tracker';
import { Session } from 'meteor/session';

Office365 = {};

Office365.requestCredential = function(options, credentialRequestCompleteCallback) {

  Session.set("redirectUrl", options.redirectUrl );

  if (!credentialRequestCompleteCallback && typeof options === 'function') {
    credentialRequestCompleteCallback = options;
    options = {};
  }

  const config = ServiceConfiguration.configurations.findOne({service: 'office365'});

  if (!config) {
    credentialRequestCompleteCallback && credentialRequestCompleteCallback(new ServiceConfiguration.ConfigError());
    return;
  }

  const credentialToken = Random.secret();

  if (!options) options = {};
  if (!options.requestPermissions) options.requestPermissions = config.permissions;

  const scope = options.requestPermissions || ['offline_access', 'user.read'];
  const flatScope = _.map(scope, encodeURIComponent).join('%20');

  const loginStyle = OAuth._loginStyle('office365', config, options);

  // The Microsoft Office 365 Application not allow the parameter "close" at redirect URLs
  const redirectUri = `${Meteor.absoluteUrl()}api/office365-auth`; //OAuth._redirectUri('office365', config).replace('?close', '');

  const loginUrl = `https://login.microsoftonline.com/${ config.tenant || 'common' }/oauth2/v2.0/authorize?client_id=${ config.clientId }&response_type=code&redirect_uri=${ redirectUri }&response_mode=query&scope=${ flatScope }&state=${ OAuth._stateParam(loginStyle, credentialToken, redirectUri) }`;

  OAuth.launchLogin({
    loginService: 'office365',
    loginStyle,
    loginUrl,
    credentialRequestCompleteCallback,
    credentialToken
  });
};