"use strict";
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
Object.defineProperty(exports, "__esModule", { value: true });
exports.Selection = void 0;
/**
 * This namespace contains functions for manipulating sets of projects
 */
class Selection {
    /**
     * Computes the intersection of two or more sets.
     */
    static intersection(first, ...rest) {
        return new Set(generateIntersection(first, ...rest));
    }
    /**
     * Computes the union of two or more sets.
     */
    static union(...sets) {
        return new Set(generateConcatenation(...sets));
    }
    /**
     * Computes a set that contains the input projects and all the direct and indirect dependencies thereof.
     */
    static expandAllDependencies(input) {
        return expandAll(input, expandDependenciesStep);
    }
    /**
     * Computes a set that contains the input projects and all projects that directly or indirectly depend on them.
     */
    static expandAllConsumers(input) {
        return expandAll(input, expandConsumers);
    }
    /**
     * Iterates the direct dependencies of the listed projects. May contain duplicates.
     */
    static *directDependenciesOf(input) {
        for (const item of input) {
            yield* item.dependencyProjects;
        }
    }
    /**
     * Iterates the projects that declare any of the listed projects as a dependency. May contain duplicates.
     */
    static *directConsumersOf(input) {
        for (const item of input) {
            yield* item.consumingProjects;
        }
    }
}
exports.Selection = Selection;
function* generateIntersection(first, ...rest) {
    for (const item of first) {
        if (rest.every((set) => set.has(item))) {
            yield item;
        }
    }
}
function* generateConcatenation(...sets) {
    for (const set of sets) {
        yield* set;
    }
}
/**
 * Adds all dependencies of the specified project to the target set.
 */
function expandDependenciesStep(project, targetSet) {
    for (const dep of project.dependencyProjects) {
        targetSet.add(dep);
    }
}
/**
 * Adds all projects that declare the specified project as a dependency to the target set.
 */
function expandConsumers(project, targetSet) {
    for (const dep of project.consumingProjects) {
        targetSet.add(dep);
    }
}
/**
 * Computes a set derived from the input by cloning it, then iterating over every member of the new set and
 * calling a step function that may add more elements to the set.
 */
function expandAll(input, expandStep) {
    const result = new Set(input);
    for (const item of result) {
        expandStep(item, result);
    }
    return result;
}
//# sourceMappingURL=Selection.js.map