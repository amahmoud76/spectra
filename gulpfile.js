"use strict";

const build = require("@microsoft/sp-build-web");

build.addSuppression(
  `Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`,
);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set("serve", result.get("serve-deprecated"));

  return result;
};

// Transpile pdfjs-dist through babel so private class fields and
// other modern syntax are downlevelled for SPFx's webpack/acorn parser.
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    if (!generatedConfiguration.module) {
      generatedConfiguration.module = { rules: [] };
    }
    generatedConfiguration.module.rules.push({
      test: /pdfjs-dist[\\/]build[\\/]pdf\.js$/,
      use: {
        loader: "babel-loader",
        options: {
          presets: [["@babel/preset-env", { targets: { ie: "11" } }]],
          plugins: [
            "@babel/plugin-proposal-class-properties",
            "@babel/plugin-proposal-private-methods",
            "@babel/plugin-proposal-private-property-in-object",
          ],
        },
      },
    });
    return generatedConfiguration;
  },
});

build.initialize(require("gulp"));
