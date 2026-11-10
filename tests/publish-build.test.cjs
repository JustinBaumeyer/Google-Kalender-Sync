const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const {test} = require('node:test');
const {build} = require('../scripts/build.cjs');
const {publish} = require('../scripts/publish-build.cjs');
const root = path.join(__dirname, '..');

function fixture(t) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'dienstplan-release-'));
  t.after(() => fs.rmSync(directory, {recursive: true, force: true}));
  const work = path.join(directory, 'work'), remote = path.join(directory, 'origin.git');
  fs.mkdirSync(work);
  const git = (...args) => execFileSync('git', args, {cwd: work, encoding:'utf8', stdio:['ignore','pipe','pipe']}).trim();
  git('init', '--bare', '--initial-branch=main', remote);
  git('init', '--initial-branch=main');
  git('config', 'user.name', 'Build test');git('config', 'user.email', 'build@example.com');
  fs.cpSync(path.join(root, 'src'), path.join(work, 'src'), {recursive:true});
  fs.cpSync(path.join(root, 'vendor'), path.join(work, 'vendor'), {recursive:true});
  fs.copyFileSync(path.join(root, '.gitignore'), path.join(work, '.gitignore'));
  fs.writeFileSync(path.join(work, 'README.md'), 'Source repository\n');
  git('add', '.');git('commit', '-m', 'Source');
  const source = git('rev-parse', 'HEAD');
  git('remote', 'add', 'origin', remote);git('push', 'origin', 'HEAD:refs/heads/main');
  const outputs = build(work);
  for (const [name, content] of outputs) {
    const file = path.join(work, name);fs.mkdirSync(path.dirname(file), {recursive:true});fs.writeFileSync(file, content);
  }
  return {work, git, source, outputs};
}

test('publishing creates one complete release commit and leaves main at its source commit', t => {
  const {work, git, source, outputs} = fixture(t);
  assert.equal(publish(work, source), true);
  assert.equal(git('ls-remote','origin','refs/heads/main').split(/\s/)[0], source);
  const release = git('ls-remote','origin','refs/heads/release').split(/\s/)[0];
  assert.equal(git('rev-parse', release+'^'), source);
  for (const [name, content] of outputs) assert.equal(git('show', release+':'+name)+'\n', content);
  assert.deepEqual(git('diff-tree','--no-commit-id','--name-only','-r',release).split('\n').sort(), [...outputs.keys()].sort());
});

test('a completed older build cannot replace the release after main has advanced', t => {
  const {work, git, source} = fixture(t);
  git('push','origin',source+':refs/heads/release');
  fs.appendFileSync(path.join(work,'README.md'),'Newer source\n');
  git('add','README.md');git('commit','-m','Newer source');git('push','origin','HEAD:refs/heads/main');
  git('checkout','--detach',source);
  assert.equal(publish(work, source),false);
  assert.equal(git('ls-remote','origin','refs/heads/release').split(/\s/)[0],source);
});

test('mismatched artifacts and unrelated staged files never get published', t => {
  const {work,git,source,outputs} = fixture(t);
  const file = path.join(work,'Code.gs');fs.appendFileSync(file,'\n// changed after tests\n');
  assert.throws(()=>publish(work,source),/artifact differs/);
  fs.writeFileSync(file,outputs.get('Code.gs'));
  fs.writeFileSync(path.join(work,'unexpected.txt'),'Do not publish');git('add','unexpected.txt');
  assert.throws(()=>publish(work,source),/Unexpected staged files/);
  assert.equal(git('ls-remote','origin','refs/heads/release'),'');
});
