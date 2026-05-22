import { mkdir, writeFile } from "node:fs/promises";

const buildId = process.env.COMMIT_REF || process.env.DEPLOY_ID || String(Date.now());
const payload = {
  buildId,
  builtAt: new Date().toISOString(),
};

await mkdir("public", { recursive: true });
await writeFile("public/app-version.json", `${JSON.stringify(payload, null, 2)}\n`);
