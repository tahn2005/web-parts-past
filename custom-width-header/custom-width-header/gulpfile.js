'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

/**
 * Keep everything the property pane needs in its one named chunk
 * ("<web part>-property-pane") instead of letting webpack split shared
 * libraries into extra numbered files. Readers still only load the main bundle.
 */
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    generatedConfiguration.optimization = generatedConfiguration.optimization || {};
    generatedConfiguration.optimization.splitChunks = false;
    return generatedConfiguration;
  }
});

build.initialize(require('gulp'));
