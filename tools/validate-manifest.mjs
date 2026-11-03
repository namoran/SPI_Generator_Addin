/* Runs tools/validate_manifest.py (Python's XML parser) on both
 * manifests. Usage: npm run validate */
import { spawnSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const res = spawnSync('python3', [path.join(root, 'tools', 'validate_manifest.py')], {
  cwd: root,
  encoding: 'utf8',
});
process.stdout.write(res.stdout || '');
process.stderr.write(res.stderr || '');
process.exit(res.status === null ? 2 : res.status);
