import { chromium, expect, Page } from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { mergeReport } from './merge-only';

const ROOT = path.dirname(fileURLToPath(import.meta.url));

const PROFILE_DIR =
  process.env.ZOHO_PROFILE_DIR ??
  path.resolve(ROOT, 'playwright/.auth/zoho-profile');

const DOWNLOAD_DIR =
  process.env.ZOHO_DOWNLOAD_DIR ??
  path.resolve(ROOT, 'downloads/zoho-daily');

const reports = [
  {
    name: 'TicketToday_Support',
    url: 'https://desk.zoho.com/agent/carmensoftware/carmen-software-support/reports/details/483929000029190842',
    output: 'TicketToday_Support.xls',
  },
  {
    name: 'OpenAllNew',
    url: 'https://desk.zoho.com/agent/carmensoftware/carmen-software-support/reports/details/483929000037008035',
    output: 'OpenAllNew.xls',
  },
];

async function openExportMenu(page: Page) {
  const notNow = page.getByText('Not Now', { exact: true });
  if (await notNow.isVisible()) await notNow.click();
  const moreButton = page.locator('button[data-id="moreIcon"]');
  await expect(moreButton).toBeEnabled({ timeout: 60_000 });
  await moreButton.click();
  await expect(page.getByText(/^Excel$/i)).toBeVisible({ timeout: 10_000 });
}

async function exportReport(page: Page, report: typeof reports[number], runDir: string) {
  console.log(`Exporting: ${report.name}`);

  await page.goto(report.url, { waitUntil: 'domcontentloaded' });

  console.log('If Zoho asks you to sign in, complete login in the browser (up to 5 minutes).');
  // Login can redirect to the home page rather than back to the report.
  await page.waitForURL(url => url.hostname === 'desk.zoho.com' && url.pathname.includes('/agent/'), { timeout: 300_000 });
  if (!page.url().includes(new URL(report.url).pathname)) {
    await page.goto(report.url, { waitUntil: 'domcontentloaded' });
  }
  await expect(page.getByText(report.name, { exact: true }).first()).toBeVisible({ timeout: 60_000 });

  await openExportMenu(page);

  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 120_000 }),
    page.getByText(/^Excel$/i).click(),
  ]);
  const failure = await download.failure();
  if (failure) throw new Error(`Download failed for ${report.name}: ${failure}`);
  const targetPath = path.join(runDir, report.output);

  await download.saveAs(targetPath);

  if (fs.statSync(targetPath).size === 0) throw new Error(`Empty export: ${targetPath}`);
  console.log(`Downloaded: ${targetPath}`);
}

async function main() {
  // A separate run folder prevents old or partial downloads entering today's report.
  fs.mkdirSync(DOWNLOAD_DIR, { recursive: true });
  const runDir = fs.mkdtempSync(path.join(DOWNLOAD_DIR, 'run-'));
  console.log(`Downloads: ${runDir}`);

  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    acceptDownloads: true,
    viewport: { width: 1600, height: 900 },
  });

  const page = context.pages()[0] ?? await context.newPage();

  try {
    for (const report of reports) {
      await exportReport(page, report, runDir);
    }
  } catch (error) {
    await page.screenshot({ path: path.join(runDir, 'download-error.png'), fullPage: true }).catch(() => {});
    console.error(`Download failed. Diagnostic screenshot and completed files: ${runDir}`);
    throw error;
  } finally {
    await context.close();
  }

  await mergeReport(runDir, {
    today: path.join(runDir, reports[0].output),
    openAll: path.join(runDir, reports[1].output),
  });

  console.log('Done.');
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});