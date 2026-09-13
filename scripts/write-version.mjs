import { mkdir, writeFile } from "node:fs/promises";

const version = "Update 1.2";
const changes = [
  "Added conflict pairs in Chemistry so selected students can be kept apart.",
  "Added Flip chemistry in Settings to reverse how chemistry scores guide seating.",
  "Changed the update notification to open a version 1.2 details pop-up.",
  "Added privacy notices on the login screen and signed-in sidebar.",
  "Kept existing student metadata, chemistry, sit-together frequency, layouts, and teacher settings intact.",
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
