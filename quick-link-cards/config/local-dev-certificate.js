// Lets `npm run start` work on computers where you don't have admin rights.
//
// By default the SPFx dev server creates a certificate and tries to add it to the
// computer's trusted certificates, which needs admin (sudo) rights. Setting
// `devServer.server` here skips that step; the dev server makes its own
// self-signed certificate instead, which you accept once in your browser by opening
// https://localhost:4321/temp/build/manifests.js and choosing to continue.
module.exports = function (config) {
  const configs = Array.isArray(config) ? config : [config];
  for (const c of configs) {
    c.devServer = { ...(c.devServer || {}), server: { type: 'https' } };
  }
  return config;
};
