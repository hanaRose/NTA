"use strict";
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
Object.defineProperty(exports, "__esModule", { value: true });
exports.PackageLookup = void 0;
class PackageLookup {
    constructor() {
        this._packageMap = new Map();
    }
    loadTree(root) {
        const queue = [root];
        // We want the lookup to return the shallowest match, so this is a breadth first
        // traversal
        for (;;) {
            const current = queue.shift();
            if (!current) {
                break;
            }
            for (const child of current.children) {
                queue.push(child);
            }
            const key = current.nameAndVersion;
            if (!this._packageMap.has(key)) {
                this._packageMap.set(key, current);
            }
        }
    }
    getPackage(nameAndVersion) {
        return this._packageMap.get(nameAndVersion);
    }
}
exports.PackageLookup = PackageLookup;
//# sourceMappingURL=PackageLookup.js.map