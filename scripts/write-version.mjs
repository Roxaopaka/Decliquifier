import { mkdir, writeFile } from "node:fs/promises";

const version = "Update 1.1";
const changes = [
  "Changed update notices to list the specific release changes.",
  "Added saved randomized seating charts in the Randomize tab.",
  "Fixed Randomize capacity so it matches the visible layout preview.",
  "Fixed hidden/out-of-room desks taking student assignments.",
  "Changed extra-seat randomization so open seats rotate across repeated runs.",
];
const buildId = process.env.COMMIT_REF || process.env.DEPLOY_ID || String(Date.now());
const payload = {
  buildId,
  version,
  changes,
  builtAt: new Date().toISOString(),
};

await mkdir("public", { recursive: true });
await writeFile("public/app-version.json", `${JSON.stringify(payload, null, 2)}\n`);
