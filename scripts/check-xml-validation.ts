import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import {
  compileSchema,
  readSchemaGraph,
  schemaFiles,
  validateXml,
} from '../test/lib/schema-validator.ts';

const repository = fileURLToPath(new URL('../', import.meta.url));
const submodule = 'references/ecma-376-5th';
const gitlink = execFileSync('git', ['ls-tree', 'HEAD', '--', submodule], {
  cwd: repository,
  encoding: 'utf8',
}).trim();
const expectedCommit = /^160000 commit ([0-9a-f]{40})\t/.exec(gitlink)?.[1];
const actualCommit = execFileSync('git', ['-C', submodule, 'rev-parse', 'HEAD'], {
  cwd: repository,
  encoding: 'utf8',
}).trim();
const changes = execFileSync(
  'git',
  ['-C', submodule, 'status', '--porcelain', '--untracked-files=all'],
  {
    cwd: repository,
    encoding: 'utf8',
  },
).trim();
if (!expectedCommit || actualCommit !== expectedCommit || changes.length > 0) {
  throw new Error('XML schema infrastructure requires the clean, exact pinned ECMA-376 submodule');
}

const documents = readSchemaGraph(Object.values(schemaFiles));
for (const [kind, url] of Object.entries(schemaFiles)) {
  const validator = compileSchema(url, documents);
  try {
    if (kind === 'rels') {
      validateXml(
        validator,
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
      );
    }
    let rejected = false;
    try {
      validateXml(validator, '<invalid-root/>');
    } catch {
      rejected = true;
    }
    if (!rejected) throw new Error(`${kind}: invalid XML was accepted`);
  } finally {
    validator.dispose();
  }
}
console.log(
  `XML validation ready: libxml2 WASM; ${Object.keys(schemaFiles).length} compiled schema families; ${Object.keys(documents).length} local XSD files; ECMA ${actualCommit}`,
);
