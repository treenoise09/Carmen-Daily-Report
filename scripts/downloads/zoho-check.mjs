// zoho-daily-report.ts
import { chromium, expect } from "@playwright/test";
import fs2 from "node:fs";
import path2 from "node:path";
import { fileURLToPath as fileURLToPath2 } from "node:url";

// merge-only.ts
import fs from "node:fs";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
var SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
var DAILY_REPORT_DIR = process.env.DAILY_REPORT_DIR ?? "C:\\Users\\media\\OneDrive\\Desktop\\Daily Report";
var INPUT_DIR = process.env.INPUT_DIR ?? path.resolve(SCRIPT_DIR, "input");
function findTemplateFile() {
  if (process.env.DAILY_REPORT_TEMPLATE) {
    return process.env.DAILY_REPORT_TEMPLATE;
  }
  if (!fs.existsSync(DAILY_REPORT_DIR)) {
    throw new Error(`Daily Report folder not found: ${DAILY_REPORT_DIR}`);
  }
  const candidates = fs.readdirSync(DAILY_REPORT_DIR).filter(
    (name) => /^Tickets_AllZoho.*\.xls$/i.test(name) && !name.startsWith("~$")
  ).map((name) => {
    const fullPath = path.join(DAILY_REPORT_DIR, name);
    const stat = fs.statSync(fullPath);
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
    const child = spawn("powershell.exe", args, {
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
async function mergeReport(inputDir = INPUT_DIR, exports) {
  const ps1Path = path.resolve(SCRIPT_DIR, "update-zoho-daily-report.ps1");
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
    ...exports ? ["-TodayExport", exports.today, "-OpenAllExport", exports.openAll] : []
  ]);
  console.log("Done.");
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) mergeReport().catch((error) => {
  console.error(error);
  process.exit(1);
});

// zoho-daily-report.ts
var ROOT = path2.dirname(fileURLToPath2(import.meta.url));
var PROFILE_DIR = process.env.ZOHO_PROFILE_DIR ?? path2.resolve(ROOT, "playwright/.auth/zoho-profile");
var DOWNLOAD_DIR = process.env.ZOHO_DOWNLOAD_DIR ?? path2.resolve(ROOT, "downloads/zoho-daily");
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
  await expect(page.getByText(/EXPORT AS/i)).toBeVisible({ timeout: 1e4 });
}
async function exportReport(page, report, runDir) {
  console.log(`Exporting: ${report.name}`);
  await page.goto(report.url, { waitUntil: "domcontentloaded" });
  console.log("If Zoho asks you to sign in, complete login in the browser (up to 5 minutes).");
  await page.waitForURL((url) => url.hostname === "desk.zoho.com" && url.pathname.includes("/agent/"), { timeout: 3e5 });
  if (!page.url().includes(new URL(report.url).pathname)) {
    await page.goto(report.url, { waitUntil: "domcontentloaded" });
  }
  await expect(page.getByText(report.name, { exact: true }).first()).toBeVisible({ timeout: 6e4 });
  await openExportMenu(page);
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 12e4 }),
    page.getByText(/^Excel$/i).click()
  ]);
  const failure = await download.failure();
  if (failure) throw new Error(`Download failed for ${report.name}: ${failure}`);
  const targetPath = path2.join(runDir, report.output);
  await download.saveAs(targetPath);
  if (fs2.statSync(targetPath).size === 0) throw new Error(`Empty export: ${targetPath}`);
  console.log(`Downloaded: ${targetPath}`);
}
async function main() {
  fs2.mkdirSync(DOWNLOAD_DIR, { recursive: true });
  const runDir = fs2.mkdtempSync(path2.join(DOWNLOAD_DIR, "run-"));
  console.log(`Downloads: ${runDir}`);
  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
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
    await page.screenshot({ path: path2.join(runDir, "download-error.png"), fullPage: true }).catch(() => {
    });
    console.error(`Download failed. Diagnostic screenshot and completed files: ${runDir}`);
    throw error;
  } finally {
    await context.close();
  }
  await mergeReport(runDir, {
    today: path2.join(runDir, reports[0].output),
    openAll: path2.join(runDir, reports[1].output)
  });
  console.log("Done.");
}
main().catch((error) => {
  console.error(error);
  process.exit(1);
});
