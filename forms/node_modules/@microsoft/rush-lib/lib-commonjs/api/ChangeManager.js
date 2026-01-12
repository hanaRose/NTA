"use strict";
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
Object.defineProperty(exports, "__esModule", { value: true });
exports.ChangeManager = void 0;
const ChangeFile_1 = require("./ChangeFile");
/**
 * A class that helps with programmatically interacting with Rush's change files.
 * @public
 */
class ChangeManager {
    /**
     * Creates a change file that has a 'none' type.
     * @param rushConfiguration - The rush configuration we are working with
     * @param projectName - The name of the project for which to create a change file
     * @param emailAddress - The email address which should be associated with this change
     * @returns the path to the file that was created, or undefined if no file was written
     */
    static createEmptyChangeFiles(rushConfiguration, projectName, emailAddress) {
        const projectInfo = rushConfiguration.getProjectByName(projectName);
        if (projectInfo && projectInfo.shouldPublish) {
            const changefile = {
                // eslint-disable-line @typescript-eslint/no-explicit-any
                changes: [
                    {
                        comment: '',
                        packageName: projectName,
                        type: 'none'
                    }
                ],
                packageName: projectName,
                email: emailAddress
            };
            return new ChangeFile_1.ChangeFile(changefile, rushConfiguration).writeSync();
        }
        return undefined;
    }
}
exports.ChangeManager = ChangeManager;
//# sourceMappingURL=ChangeManager.js.map