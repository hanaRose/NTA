export interface IWorkspaceFileSaveOptions {
    /**
     * If there is an existing file, and the contents have not changed, then
     * don't write anything; this preserves the old timestamp.
     */
    onlyIfChanged?: boolean;
    /**
     * Creates the folder recursively using FileSystem.ensureFolder()
     * Defaults to false.
     */
    ensureFolderExists?: boolean;
}
/**
 * This class is a parser for pnpm's pnpm-workspace.yaml file format.
 */
export declare abstract class BaseWorkspaceFile {
    protected _alreadyWarnedSpecs: Set<string>;
    /**
     * Serializes and saves the workspace file to specified location
     */
    save(filePath: string, options: IWorkspaceFileSaveOptions): void;
    /**
     * Adds a package path to the workspace file.
     *
     * @virtual
     */
    abstract addPackage(packagePath: string): void;
    /** @virtual */
    protected abstract serialize(): string;
}
//# sourceMappingURL=BaseWorkspaceFile.d.ts.map