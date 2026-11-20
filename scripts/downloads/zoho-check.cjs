var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));

// zoho-daily-report.ts
var import_test = require("@playwright/test");
var import_node_fs2 = __toESM(require("node:fs"));
var import_node_path2 = __toESM(require("node:path"));
var import_node_url2 = require("node:url");

// merge-only.ts
var import_node_fs = __toESM(require("node:fs"));
var import_node_path = __toESM(require("node:path"));
var import_node_child_process = require("node:child_process");
var import_node_url = require("node:url");
var import_meta = {};
var SCRIPT_DIR = import_node_path.default.dirname((0, import_node_url.fileURLToPath)(import_meta.url));
var DAILY_REPORT_DIR = process.env.DAILY_REPORT_DIR ?? "C:\\Users\\media\\OneDrive\\Desktop\\Daily Report";
var INPUT_DIR = process.env.INPUT_DIR ?? import_node_path.default.resolve(SCRIPT_DIR, "input");
function findTemplateFile() {
  if (process.env.DAILY_REPORT_TEMPLATE) {
    return process.env.DAILY_REPORT_TEMPLATE;
  }
  if (!import_node_fs.default.existsSync(DAILY_REPORT_DIR)) {
    throw new Error(`Daily Report folder not found: ${DAILY_REPORT_DIR}`);
  }
  const candidates = import_node_fs.default.readdirSync(DAILY_REPORT_DIR).filter(
    (name) => /^Tickets_AllZoho.*\.xls$/i.test(name) && !name.startsWith("~$")
  ).map((name) => {
    const fullPath = import_node_path.default.join(DAILY_REPORT_DIR, name);
    const stat = import_node_fs.default.statSync(fullPath);
    return {
      name,
      fullPath,
      mtimeMs: stat.mtimeMs
    };
  }).sort((a, b) => b.mtimeMs - a.mtimeMs);
  if (candidates.length === 0) {
    throw new Error(
      [
        `Cannot find template file in: ${DAILY_REPORT_DIR}`,
        `Expected something like: Tickets_AllZoho*.xls`,
        `Put your old working report file in that folder, or set DAILY_REPORT_TEMPLATE manually.`
      ].join("\n")
    );
  }
  return candidates[0].fullPath;
}
function clearLine() {
  process.stdout.write("\r\x1B[K");
}
function runPowerShellWithLoading(args) {
  return new Promise((resolve, reject) => {
    const frames = ["\u280B", "\u2819", "\u2839", "\u2838", "\u283C", "\u2834", "\u2826", "\u2827", "\u2807", "\u280F"];
    let frameIndex = 0;
    const startedAt = Date.now();
    const child = (0, import_node_child_process.spawn)("powershell.exe", args, {
      windowsHide: true,
      stdio: ["ignore", "pipe", "pipe"]
    });
    const timer = setInterval(() => {
      const elapsedSeconds = Math.floor((Date.now() - startedAt) / 1e3);
      const frame = frames[frameIndex % frames.length];
      frameIndex++;
      clearLine();
      process.stdout.write(`${frame} Generating daily Zoho report... ${elapsedSeconds}s`);
    }, 120);
    child.stdout.on("data", (data) => {
      clearLine();
      process.stdout.write(data.toString());
    });
    child.stderr.on("data", (data) => {
      clearLine();
      process.stderr.write(data.toString());
    });
    child.on("error", (error) => {
      clearInterval(timer);
      clearLine();
      reject(error);
    });
    child.on("close", (code) => {
      clearInterval(timer);
      clearLine();
      if (code === 0) {
        resolve();
      } else {
        reject(new Error(`PowerShell exited with code ${code}`));
      }
    });
  });
}
async function mergeReport(inputDir = INPUT_DIR, exports2) {
  const ps1Path = import_node_path.default.resolve(SCRIPT_DIR, "update-zoho-daily-report.ps1");
  const templatePath = findTemplateFile();
  console.log(`Using template: ${templatePath}`);
  console.log(`Using input dir: ${inputDir}`);
  console.log(`Using output dir: ${DAILY_REPORT_DIR}`);
  console.log("");
  await runPowerShellWithLoading([
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    ps1Path,
    "-TemplatePath",
    templatePath,
    "-InputDir",
    inputDir,
    "-OutputDir",
    DAILY_REPORT_DIR,
    ...exports2 ? ["-TodayExport", exports2.today, "-OpenAllExport", exports2.openAll] : []
  ]);
  console.log("Done.");
}
if (process.argv[1] && import_node_path.default.resolve(process.argv[1]) === (0, import_node_url.fileURLToPath)(import_meta.url)) mergeReport().catch((error) => {
  console.error(error);
  process.exit(1);
});

// zoho-daily-report.ts
var import_meta2 = {};
var ROOT = import_node_path2.default.dirname((0, import_node_url2.fileURLToPath)(import_meta2.url));
var PROFILE_DIR = process.env.ZOHO_PROFILE_DIR ?? import_node_path2.default.resolve(ROOT, "playwright/.auth/zoho-profile");
var DOWNLOAD_DIR = process.env.ZOHO_DOWNLOAD_DIR ?? import_node_path2.default.resolve(ROOT, "downloads/zoho-daily");
var reports = [
  {
    name: "TicketToday_Support",
    url: "https://desk.zoho.com/agent/carmensoftware/carmen-software-support/reports/details/483929000029190842",
    output: "TicketToday_Support.xls"
  },
  {
    name: "OpenAllNew",
    url: "https://desk.zoho.com/agent/carmensoftware/carmen-software-support/reports/details/483929000037008035",
    output: "OpenAllNew.xls"
  }
];
async function firstVisible(candidates, timeout = 1500) {
  for (const locator of candidates) {
    try {
      const first = locator.first();
      await first.waitFor({ state: "visible", timeout });
      return first;
    } catch {
    }
  }
  return null;
}
async function openExportMenu(page) {
  const moreButton = await firstVisible([
    page.getByRole("button", { name: /more|more actions|options|ellipsis/i }),
    page.locator('[aria-label*="More" i]'),
    page.locator('[title*="More" i]')
  ]);
  if (moreButton) {
    await moreButton.click();
  } else {
    const createChart = page.getByRole("button", { name: /Create Chart/i });
    const box = await createChart.boundingBox();
    if (!box) {
      throw new Error("Cannot find Zoho more/export button. Update openExportMenu selector.");
    }
    await page.mouse.click(box.x + box.width + 55, box.y + box.height / 2);
  }
  await (0, import_test.expect)(page.getByText(/EXPORT AS/i)).toBeVisible({ timeout: 1e4 });
}
async function exportReport(page, report, runDir) {
  console.log(`Exporting: ${report.name}`);
  await page.goto(report.url, { waitUntil: "domcontentloaded" });
  console.log("If Zoho asks you to sign in, complete login in the browser (up to 5 minutes).");
  await page.waitForURL((url) => url.hostname === "desk.zoho.com" && url.pathname.includes("/agent/"), { timeout: 3e5 });
  if (!page.url().includes(new URL(report.url).pathname)) {
    await page.goto(report.url, { waitUntil: "domcontentloaded" });
  }
  await (0, import_test.expect)(page.getByText(report.name, { exact: true }).first()).toBeVisible({ timeout: 6e4 });
  await openExportMenu(page);
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 12e4 }),
    page.getByText(/^Excel$/i).click()
  ]);
  const failure = await download.failure();
  if (failure) throw new Error(`Download failed for ${report.name}: ${failure}`);
  const targetPath = import_node_path2.default.join(runDir, report.output);
  await download.saveAs(targetPath);
  if (import_node_fs2.default.statSync(targetPath).size === 0) throw new Error(`Empty export: ${targetPath}`);
  console.log(`Downloaded: ${targetPath}`);
}
async function main() {
  import_node_fs2.default.mkdirSync(DOWNLOAD_DIR, { recursive: true });
  const runDir = import_node_fs2.default.mkdtempSync(import_node_path2.default.join(DOWNLOAD_DIR, "run-"));
  console.log(`Downloads: ${runDir}`);
  const context = await import_test.chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    acceptDownloads: true,
    viewport: { width: 1600, height: 900 }
  });
  const page = context.pages()[0] ?? await context.newPage();
  try {
    for (const report of reports) {
      await exportReport(page, report, runDir);
    }
  } catch (error) {
    await page.screenshot({ path: import_node_path2.default.join(runDir, "download-error.png"), fullPage: true }).catch(() => {
    });
    console.error(`Download failed. Diagnostic screenshot and completed files: ${runDir}`);
    throw error;
  } finally {
    await context.close();
  }
  await mergeReport(runDir, {
    today: import_node_path2.default.join(runDir, reports[0].output),
    openAll: import_node_path2.default.join(runDir, reports[1].output)
  });
  console.log("Done.");
}
main().catch((error) => {
  console.error(error);
  process.exit(1);
});
