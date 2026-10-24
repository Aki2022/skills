// RFC 8785: ECMAScript primitives and recursive UTF-16 property ordering.
// Emit object members directly: JSON.stringify(object) reorders integer-looking keys.
'use strict';
function canonicalize(value) {
  if (value === null || typeof value === 'boolean') return JSON.stringify(value);
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error('non-finite number');
    return JSON.stringify(value);
  }
  if (typeof value === 'string') {
    for (const character of value) {
      const code = character.codePointAt(0);
      if (code >= 0xd800 && code <= 0xdfff) throw new Error('invalid Unicode');
    }
    return JSON.stringify(value);
  }
  if (Array.isArray(value)) return '[' + value.map(canonicalize).join(',') + ']';
  if (value && typeof value === 'object' && Object.getPrototypeOf(value) === Object.prototype) {
    return '{' + Object.keys(value).sort().map(key => canonicalize(key) + ':' + canonicalize(value[key])).join(',') + '}';
  }
  throw new Error('not I-JSON data');
}
module.exports = { canonicalize };
if (require.main === module) {
  try {
    const values = JSON.parse(require('node:fs').readFileSync(0, 'utf8'));
    if (!Array.isArray(values)) throw new Error('batch must be an array');
    process.stdout.write(JSON.stringify(values.map(canonicalize)));
  } catch (_) {
    process.stderr.write('JCS input is invalid\n');
    process.exitCode = 2;
  }
}
