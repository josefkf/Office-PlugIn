// Usage:
//   npm run build                                   -> dist/ for https://localhost:3000
//   BASE_URL=https://assets.contoso.com npm run build -> dist/ for production
import * as esbuild from "esbuild";
import { cpSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";

const BASE_URL = (process.env.BASE_URL || "https://localhost:3000").replace(/\/$/, "");
const PROVIDER = process.env.PROVIDER || "Contoso IT";
const MOCK_AUTH = process.argv.includes("--mock-auth"); // test builds only

rmSync("dist", { recursive: true, force: true });
mkdirSync("dist", { recursive: true });

const mockAuthPlugin = {
  name: "mock-auth",
  setup(build) {
    build.onResolve({ filter: /\.\/auth\.js$/ }, () => ({ path: new URL("./test/mock-auth.js", import.meta.url).pathname }));
  },
};

await esbuild.build({
  entryPoints: ["src/taskpane.js"],
  bundle: true,
  minify: !MOCK_AUTH,
  sourcemap: MOCK_AUTH,
  format: "iife",
  target: ["es2020"],
  outfile: "dist/taskpane.bundle.js",
  plugins: MOCK_AUTH ? [mockAuthPlugin] : [],
});

for (const f of ["taskpane.html", "taskpane.css", "commands.html"]) cpSync(`src/${f}`, `dist/${f}`);
cpSync("assets", "dist/assets", { recursive: true });

const manifest = readFileSync("manifest.template.xml", "utf8")
  .replaceAll("{{BASE_URL}}", BASE_URL)
  .replaceAll("{{PROVIDER}}", PROVIDER);
writeFileSync("dist/manifest.xml", manifest);

console.log(`Built dist/ for ${BASE_URL}${MOCK_AUTH ? " (mock auth)" : ""}`);
