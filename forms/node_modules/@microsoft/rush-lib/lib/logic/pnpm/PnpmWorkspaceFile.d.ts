import { BaseWorkspaceFile } from '../base/BaseWorkspaceFile';
export declare class PnpmWorkspaceFile extends BaseWorkspaceFile {
    /**
     * The filename of the workspace file.
     */
    readonly workspaceFilename: string;
    private _workspacePackages;
    /**
     * The PNPM workspace file is used to specify the location of workspaces relative to the root
     * of your PNPM install.
     */
    constructor(workspaceYamlFilename: string);
    /** @override */
    addPackage(packagePath: string): void;
    /** @override */
    protected serialize(): string;
}
//# sourceMappingURL=PnpmWorkspaceFile.d.ts.map