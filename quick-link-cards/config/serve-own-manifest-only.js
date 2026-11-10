// Fixes "Could not load quick-link-cards-web-part in require. TypeError: Cannot
// destructure property 'id'..." when testing with `npm run start`.
//
// By default the dev server's manifests.js also lists SharePoint's own framework
// libraries (sp-core-library, sp-webpart-base, ...) from node_modules. The page then
// prefers those local copies over the versions SharePoint Online is actually running,
// and they reference internal SharePoint bundles that no longer exist there. Leaving
// them out makes the page use its own framework, exactly like a deployed web part.
const fs = require('fs');
const path = require('path');

function ownComponentIds(srcFolder) {
  const ids = new Set();
  const walk = folder => {
    for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
      const full = path.join(folder, entry.name);
      if (entry.isDirectory()) {
        walk(full);
      } else if (entry.name.endsWith('.manifest.json')) {
        // Manifests allow comments, so pull the id out with a regex instead of JSON.parse.
        const match = /"id"\s*:\s*"([^"]+)"/.exec(fs.readFileSync(full, 'utf8'));
        if (match) {
          ids.add(match[1]);
        }
      }
    }
  };
  walk(srcFolder);
  return ids;
}

module.exports = function (config) {
  const ownIds = ownComponentIds(path.resolve(__dirname, '..', 'src'));
  const configs = Array.isArray(config) ? config : [config];
  for (const c of configs) {
    for (const plugin of c.plugins || []) {
      const processor = plugin && plugin.constructor && plugin.constructor.name === 'ManifestPlugin'
        ? plugin._options && plugin._options.cumulativeManifestProcessor
        : undefined;
      if (!processor) {
        continue;
      }
      const referenced = processor.discoverManifests(processor._options.rootPath, 4);
      const ignore = Object.keys(referenced).filter(id => !ownIds.has(id));
      processor._options.ignoreOutputManifestIds = [
        ...(processor._options.ignoreOutputManifestIds || []),
        ...ignore
      ];
    }
  }
  return config;
};
