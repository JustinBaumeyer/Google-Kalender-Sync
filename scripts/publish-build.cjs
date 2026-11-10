const { execFileSync } = require("node:child_process");
const path = require("node:path");
const { build, check } = require("./build.cjs");

function publish(directory, sourceCommit) {
  if (!/^[a-f0-9]{40}$/.test(sourceCommit || "")) throw new Error("Invalid source commit");
  const git = (...args) =>
    execFileSync("git", args, {
      cwd: directory,
      encoding: "utf8",
      stdio: ["ignore", "pipe", "pipe"]
    }).trim();
  const remoteHead = (branch) => git("ls-remote", "origin", "refs/heads/" + branch).split(/\s/)[0];
  if (git("rev-parse", "HEAD") !== sourceCommit)
    throw new Error("Checkout differs from tested source");
  if (remoteHead("main") !== sourceCommit) {
    console.log("A newer source commit exists; this build will not be published.");
    return false;
  }
  if (check(directory).length) throw new Error("Build artifact differs from tested source");
  const previous = remoteHead("release");
  const files = [...build(directory).keys()];
  git("add", "--force", "--", ...files);
  const staged = git("diff", "--cached", "--name-only").split("\n").filter(Boolean);
  if (staged.some((name) => !files.includes(name))) throw new Error("Unexpected staged files");
  git(
    "-c",
    "user.name=github-actions[bot]",
    "-c",
    "user.email=41898282+github-actions[bot]@users.noreply.github.com",
    "commit",
    "--allow-empty",
    "-m",
    "Build Apps Script from " + sourceCommit
  );
  if (remoteHead("main") !== sourceCommit) {
    console.log("Source changed during publishing; keeping the previous release.");
    return false;
  }
  // One ref update publishes the entire tested build. Reject a concurrent edit.
  git(
    "push",
    "--force-with-lease=refs/heads/release:" + previous,
    "origin",
    "HEAD:refs/heads/release"
  );
  console.log("Published build from " + sourceCommit);
  return true;
}

if (require.main === module) publish(path.join(__dirname, ".."), process.env.GITHUB_SHA);
module.exports = { publish };
