// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * DevServer APIs index file
 * Exports manifest API and dev server components registration
 */

const manifestApi = require('./manifestApi');

/**
 * Register dev server manifest APIs with an Express application
 * @param {object} app Express application
 */
function registerDevServerApis(app) {
  console.log('*** Mounting Manifest API ***');
  app.use('/', manifestApi);
}

function registerDevServerComponents() {
  console.log("*********************************************************************");
  console.log('***                Mounting Dev Server Components                ***');

  console.log("*********************************************************************");
}

module.exports = {
  manifestApi,
  registerDevServerApis,
  registerDevServerComponents
};
