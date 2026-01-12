import type { TSESLint } from '@typescript-eslint/utils';
import { InputFileMessageIds, ImportMessageIds } from './PackletAnalyzer';
export type MessageIds = InputFileMessageIds | ImportMessageIds;
type Options = [];
declare const mechanics: TSESLint.RuleModule<MessageIds, Options>;
export { mechanics };
//# sourceMappingURL=mechanics.d.ts.map