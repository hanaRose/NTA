declare const EventEmitter: any;
declare const childProcess: any;
declare const childProcessActual: any;
declare let spawnMockConfig: {
    emitError: any;
    returnCode: any;
};
/**
 * Helper to initialize how the `spawn` mock should behave.
 */
declare function normalizeSpawnMockConfig(maybeConfig?: any): {
    emitError: any;
    returnCode: any;
};
/**
 * Initialize the `spawn` mock behavior.
 *
 * Not a pure function.
 */
declare function setSpawnMockConfig(spawnConfig: any): void;
/**
 * Mock of `spawn`.
 */
declare function spawn(file: string, args: string[], options: {}): any;
//# sourceMappingURL=child_process.d.ts.map