"use strict";
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
Object.defineProperty(exports, "__esModule", { value: true });
exports.BaseWorkspaceFile = void 0;
const node_core_library_1 = require("@rushstack/node-core-library");
/**
 * This class is a parser for pnpm's pnpm-workspace.yaml file format.
 */
class BaseWorkspaceFile {
    constructor() {
        this._alreadyWarnedSpecs = new Set();
    }
    /**
     * Serializes and saves the workspace file to specified location
     */
    save(filePath, options) {
        // Do we need to read the previous file contents?
        let oldBuffer = undefined;
        if (options.onlyIfChanged && node_core_library_1.FileSystem.exists(filePath)) {
            try {
                oldBuffer = node_core_library_1.FileSystem.readFileToBuffer(filePath);
            }
            catch (error) {
                // Ignore this error, and try writing a new file.  If that fails, then we should report that
                // error instead.
            }
        }
        const newYaml = this.serialize();
        const newBuffer = Buffer.from(newYaml); // utf8 encoding happens here
        if (options.onlyIfChanged) {
            // Has the file changed?
            if (oldBuffer && Buffer.compare(newBuffer, oldBuffer) === 0) {
                // Nothing has changed, so don't touch the file
                return;
            }
        }
        node_core_library_1.FileSystem.writeFile(filePath, newBuffer.toString(), {
            ensureFolderExists: options.ensureFolderExists
        });
    }
}
exports.BaseWorkspaceFile = BaseWorkspaceFile;
//# sourceMappingURL=BaseWorkspaceFile.js.map