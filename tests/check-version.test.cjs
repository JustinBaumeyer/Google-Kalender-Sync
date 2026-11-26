const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const {test} = require('node:test');
const {build} = require('../scripts/build.cjs');
const {checkVersion, report} = require('../scripts/check-version.cjs');
const root = path.join(__dirname, '..');

function fixture(t) {
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'dienstplan-version-'));
  t.after(() => fs.rmSync(work, {recursive: true, force: true}));
  const git = (...args) => execFileSync('git', args, {cwd: work, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe']}).trim();
  git('init', '--initial-branch=main');
  git('config', 'user.name', 'Version test');git('config', 'user.email', 'version@example.com');
  for (const folder of ['src', 'vendor', 'migration']) fs.cpSync(path.join(root, folder), path.join(work, folder), {recursive: true});
  const release = () => {
    for (const [name, content] of build(work)) {
      const file = path.join(work, name);fs.mkdirSync(path.dirname(file), {recursive: true});fs.writeFileSync(file, content);
    }
    git('add', '--all');git('commit', '--allow-empty', '-m', 'Release');git('branch', '--force', 'release');
  };
  const edit = (name, change) => fs.writeFileSync(path.join(work, name), change(fs.readFileSync(path.join(work, name), 'utf8')));
  return {work, release, edit};
}

test('a changed program without a higher version is reported before it is released', t => {
  const {work, release, edit} = fixture(t);
  assert.equal(checkVersion(work, 'release').status, 'unpublished');
  release();
  const version = checkVersion(work, 'release').current;
  assert.deepEqual(checkVersion(work, 'release'), {status: 'unchanged', current: version, released: version, changed: [], updater: []});

  edit('src/settings/preferences.js', source => source.replace('howFrequent: 15', 'howFrequent: 30'));
  const missing = checkVersion(work, 'release');
  assert.equal(missing.status, 'missing');assert.deepEqual(missing.changed, ['Code.gs']);
  assert.equal(report(missing)[0][0], 'warning');

  edit('src/version.js', source => source.replace('var version = ' + version, 'var version = ' + (version + 1)));
  const bumped = checkVersion(work, 'release');
  assert.equal(bumped.status, 'bumped');assert.equal(bumped.current, version + 1);
  assert.ok(report(bumped).every(line => line[0] === 'notice'));
});

test('page-only changes need a version as well, while the manual updater is only announced', t => {
  const {work, release, edit} = fixture(t);
  release();
  edit('src/web/styles.css', source => source + '\n.new-rule { color: red; }\n');
  assert.deepEqual(checkVersion(work, 'release').changed, ['Index.html']);
  assert.equal(checkVersion(work, 'release').status, 'missing');
  release();
  edit('src/updates/owner.js', source => source + '\n// changed\n');
  const updater = checkVersion(work, 'release');
  assert.equal(updater.status, 'unchanged');assert.deepEqual(updater.updater, ['admin/Updates.gs']);
  assert.match(report(updater).at(-1)[1], /copy it manually/);
});
