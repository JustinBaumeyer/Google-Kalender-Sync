const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {test} = require('node:test');
const espree = require('espree');
const root = path.join(__dirname, '..');
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'src/manifest.json'), 'utf8'));

function declarations(file) {
  const program = espree.parse(fs.readFileSync(path.join(root, file), 'utf8'), {ecmaVersion: 'latest'});
  return program.body.flatMap(node =>
    node.type === 'FunctionDeclaration' ? [node.id.name] :
    node.type === 'VariableDeclaration' ? node.declarations.filter(item => item.id.type === 'Identifier').map(item => item.id.name) : []);
}

// The files of one bundle share a single global namespace. A second declaration of
// the same name would silently replace the first one after concatenation.
for (const [bundle, files] of [
  ['Code.gs', manifest.scripts['Code.gs'].filter(file => file !== 'vendor/ical.js')],
  ['Index.html', manifest.page.scripts],
]) {
  test('every top-level name in ' + bundle + ' is declared in exactly one source file', () => {
    const seen = new Map(Object.keys((manifest.text || {})[bundle] || {}).map(name => [name, 'src/manifest.json']));
    for (const file of files) {
      for (const name of declarations(file)) {
        assert.equal(seen.get(name), undefined, name + ' is declared in ' + seen.get(name) + ' and ' + file);
        seen.set(name, file);
      }
    }
  });
}

test('browser code only calls server functions that exist and are public', () => {
  const server = new Set(manifest.scripts['Code.gs'].filter(file => file.startsWith('src/')).flatMap(declarations));
  const client = manifest.page.scripts.map(file => fs.readFileSync(path.join(root, file), 'utf8')).join('\n');
  const called = new Set(Array.from(client.matchAll(/\b(?:call|action)\(\s*"(\w+)"/g), match => match[1]));
  assert.ok(called.size > 10);
  for (const name of called) {
    assert.ok(server.has(name), name + ' is not defined on the server');
    assert.equal(name.endsWith('_'), false, name + ' is private and cannot be called from the page');
  }
});

test('the page keeps every element the scripts look up by ID', () => {
  const html = [manifest.page.template, ...Object.values(manifest.page.views)]
    .map(file => fs.readFileSync(path.join(root, file), 'utf8')).join('\n');
  const ids = new Set(Array.from(html.matchAll(/\sid="([\w-]+)"/g), match => match[1]));
  const client = manifest.page.scripts.map(file => fs.readFileSync(path.join(root, file), 'utf8')).join('\n');
  const missing = Array.from(new Set(Array.from(client.matchAll(/\$\("([\w-]+)"\)/g), match => match[1])))
    .filter(id => !ids.has(id));
  assert.deepEqual(missing, []);
});
