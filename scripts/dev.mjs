import { spawn } from "node:child_process";

const npm = process.platform === "win32" ? "npm.cmd" : "npm";
const viteArgs = process.argv.slice(2);
const spawnOptions = { stdio: "inherit", shell: process.platform === "win32" };

const processes = [
  spawn(npm, ["run", "api"], spawnOptions),
  spawn(npm, ["run", "vite", "--", ...viteArgs], spawnOptions),
];

function stopAll(signal) {
  for (const child of processes) {
    if (!child.killed) child.kill(signal);
  }
}

for (const child of processes) {
  child.on("exit", (code) => {
    if (code && code !== 0) {
      stopAll("SIGTERM");
      process.exitCode = code;
    }
  });
}

process.on("SIGINT", () => {
  stopAll("SIGINT");
});

process.on("SIGTERM", () => {
  stopAll("SIGTERM");
});
