// The measurement profiles and thresholds (references/measurement-profiles.json) for the Node callers. A missing or
// unknown profile throws: a measurement without a profile is not a pass.
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const REGISTRY = join(dirname(fileURLToPath(import.meta.url)), "..", "references", "measurement-profiles.json");

export function loadRegistry() {
  return JSON.parse(readFileSync(REGISTRY, "utf8"));
}

export function profileOf(registry, id) {
  if (typeof id !== "string" || !Object.hasOwn(registry.profiles, id)) throw new Error("invalid measurement profile");
  return registry.profiles[id];
}
