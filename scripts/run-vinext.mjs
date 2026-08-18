import { spawnSync } from "node:child_process";
import path from "node:path";

const command = process.argv[2];

if (!command) {
  console.error("Missing vinext subcommand.");
  process.exit(1);
}

const env = {
  ...process.env,
  WRANGLER_LOG_PATH: ".wrangler/wrangler.log",
};

const result =
  process.platform === "win32"
    ? spawnSync(
        "cmd.exe",
        [
          "/d",
          "/s",
          "/c",
          `""${path.join(process.cwd(), "node_modules", ".bin", "vinext.cmd")}" ${command}"`,
        ],
        {
          stdio: "inherit",
          shell: false,
          env,
        },
      )
    : spawnSync(path.join(process.cwd(), "node_modules", ".bin", "vinext"), [command], {
        stdio: "inherit",
        shell: false,
        env,
      });

if (result.error) {
  console.error(result.error.message);
  process.exit(1);
}

process.exit(result.status ?? 1);
