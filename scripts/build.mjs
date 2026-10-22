// Bundles the task pane into dist/. With --serve, rebuilds on change and
// serves dist/ on https://localhost:3000 for sideloading during development.
// With --site, builds into docs/ for GitHub Pages and writes a matching
// manifest-github-pages.xml, so the add-in can be installed without Node.js.
import * as esbuild from "esbuild";
import { cp, mkdir, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

const serve = process.argv.includes("--serve");
const site = process.argv.includes("--site");
const outdir = site ? "docs" : "dist";

const SITE_URL = "https://urzahil.github.io/Outlook-Monitoring-Plugin";
// Separate add-in ID so the hosted and localhost versions can be installed side by side.
const SITE_ADDIN_ID = "6f1d2c7e-3b4a-4e8f-9a51-2d7c0b9e4f13";

await mkdir(outdir, { recursive: true });
await Promise.all([
  cp("src/taskpane.html", `${outdir}/taskpane.html`),
  cp("src/taskpane.css", `${outdir}/taskpane.css`),
  cp("assets", `${outdir}/assets`, { recursive: true }),
  // Stops GitHub Pages from running the folder through Jekyll.
  site && writeFile(`${outdir}/.nojekyll`, ""),
]);

const options = {
  entryPoints: ["src/taskpane.js"],
  bundle: true,
  format: "esm",
  target: "es2020",
  outdir,
  minify: !serve,
  sourcemap: serve,
};

if (serve) {
  // Office requires HTTPS; office-addin-dev-certs installs a trusted localhost certificate.
  const { getHttpsServerOptions } = await import("office-addin-dev-certs");
  const { cert, key } = await getHttpsServerOptions();
  const certfile = join(tmpdir(), "addin-dev.crt");
  const keyfile = join(tmpdir(), "addin-dev.key");
  await writeFile(certfile, cert);
  await writeFile(keyfile, key);

  const ctx = await esbuild.context(options);
  await ctx.watch();
  await ctx.serve({ servedir: "dist", port: 3000, certfile, keyfile });
  console.log("Serving https://localhost:3000/taskpane.html");
} else {
  await esbuild.build(options);
  console.log(`Built ${outdir}/`);
}

if (site) {
  const manifest = (await readFile("manifest.xml", "utf8"))
    .replaceAll("https://localhost:3000", SITE_URL)
    .replace(/<Id>[^<]+<\/Id>/, `<Id>${SITE_ADDIN_ID}</Id>`);
  await writeFile("manifest-github-pages.xml", manifest);
  console.log("Wrote manifest-github-pages.xml");
}
