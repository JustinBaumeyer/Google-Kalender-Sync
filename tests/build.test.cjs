const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');
const {test} = require('node:test');
const {build, check} = require('../scripts/build.cjs');
const root = path.join(__dirname, '..');

test('built installation files match the source and the updater file list', () => {
  assert.deepEqual(check(), []);
  const outputs = build();
  const context = {PropertiesService: {getScriptProperties: () => ({})}};
  vm.runInNewContext(outputs.get('admin/Updates.gs'), context);
  const appFiles = [...outputs.keys()].filter(name => !name.startsWith('admin/')).sort();
  const managed = Array.from(context.managedUpdateFiles,
    name => name + (name === 'Index' ? '.html' : '.gs'));
  assert.deepEqual(appFiles, managed.sort());
  assert.ok(!outputs.has('settings.gs'));
  assert.ok(!outputs.has('appsscript.json'));
  assert.match(outputs.get('Code.gs'), /^var version = \d+;/);
});

test('source changes are caught before stale installation files can be shipped', t => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'dienstplan-build-'));
  t.after(() => fs.rmSync(directory, {recursive: true, force: true}));
  fs.cpSync(path.join(root, 'src'), path.join(directory, 'src'), {recursive: true});
  fs.cpSync(path.join(root, 'vendor'), path.join(directory, 'vendor'), {recursive: true});
  for (const [name, source] of build()) {
    const file = path.join(directory, name);
    fs.mkdirSync(path.dirname(file), {recursive: true});
    fs.writeFileSync(file, source);
  }
  assert.deepEqual(check(directory), []);
  const file = path.join(directory, 'src/settings/defaults.js');
  fs.writeFileSync(file, fs.readFileSync(file, 'utf8').replace('howFrequent: 15', 'howFrequent: 30'));
  assert.deepEqual(check(directory), ['Code.gs']);
  const properties = {getProperty: () => null};
  const context = {PropertiesService: {getScriptProperties: () => properties, getUserProperties: () => properties}};
  vm.runInNewContext(build(directory).get('Code.gs'), context);
  assert.equal(context.preferenceDefaults.howFrequent, 30);
  assert.equal(fs.readFileSync(path.join(directory, 'Code.gs'), 'utf8'), build().get('Code.gs'));

  const orphan = path.join(directory, 'src/forgotten.js');
  fs.writeFileSync(orphan, 'function forgotten_() {}');
  assert.throws(() => build(directory), /missing from src\/manifest.json/);
  fs.unlinkSync(orphan);
  fs.writeFileSync(file, 'function broken(');
  assert.throws(() => build(directory), SyntaxError);
});
