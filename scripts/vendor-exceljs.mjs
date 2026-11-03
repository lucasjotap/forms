// Copies the ExcelJS browser build (pinned devDependency) into public/vendor so the
// form can lazy-load it from the same origin instead of a third-party CDN.
import { copyFileSync, mkdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const pkgDir = join(root, "node_modules", "exceljs");
const { version, license } = JSON.parse(readFileSync(join(pkgDir, "package.json"), "utf8"));
const outDir = join(root, "public", "vendor");
mkdirSync(outDir, { recursive: true });
copyFileSync(join(pkgDir, "dist", "exceljs.min.js"), join(outDir, "exceljs.min.js"));
copyFileSync(join(pkgDir, "LICENSE"), join(outDir, "exceljs.LICENSE.txt"));
const size = statSync(join(outDir, "exceljs.min.js")).size;
console.log(`exceljs ${version} (${license}) -> public/vendor/exceljs.min.js (${size} bytes)`);
