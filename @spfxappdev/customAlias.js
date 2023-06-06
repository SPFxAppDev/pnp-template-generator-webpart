"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.resolveCustomAlias = void 0;
const path = require('path');
function resolveCustomAlias(build) {
    build.configureWebpack.mergeConfig({
        additionalConfiguration: (generatedConfiguration) => {
            if (!generatedConfiguration.resolve.alias) {
                generatedConfiguration.resolve.alias = {};
            }
            generatedConfiguration.resolve.alias['@webparts'] = path.resolve(__dirname, 'lib/webparts');
            generatedConfiguration.resolve.alias['@components'] = path.resolve(__dirname, 'lib/components');
            generatedConfiguration.resolve.alias['@src'] = path.resolve(__dirname, 'lib');
            return generatedConfiguration;
        }
    });
}
exports.resolveCustomAlias = resolveCustomAlias;
//# sourceMappingURL=customAlias.js.map