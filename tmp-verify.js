const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const result = spawnSync(process.execPath, ['--test', '--experimental-strip-types', 'src/lib/custom-summary-processor.test.ts'], {
  cwd: __dirname,
  encoding: 'utf8',
  shell: false,
});
const output = [
  result.stdout || '',
  result.stderr || '',
  `EXIT_CODE=${result.status ?? 'null'}`,
].join('\n');
fs.writeFileSync('verify-report.txt', output, 'utf8');
console.log(output);
