"use strict";
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
Object.defineProperty(exports, "__esModule", { value: true });
exports.PackageNameParsers = void 0;
const node_core_library_1 = require("@rushstack/node-core-library");
class PackageNameParsers {
}
exports.PackageNameParsers = PackageNameParsers;
/**
 * This is the default for `RushConfiguration.packageNameParser`.
 */
PackageNameParsers.rushDefault = new node_core_library_1.PackageNameParser({});
/**
 * This is the `RushConfiguration.packageNameParser` used when `allowMostlyStandardPackageNames = true`
 * in rush.json.
 */
PackageNameParsers.mostlyStandard = new node_core_library_1.PackageNameParser({
    allowUpperCase: true
});
/**
 * Use this in contexts where we don't have easy access to `RushConfiguration.packageNameParser`
 * AND the package name was already validated at some earlier stage.
 */
PackageNameParsers.permissive = PackageNameParsers.mostlyStandard;
//# sourceMappingURL=PackageNameParsers.js.map