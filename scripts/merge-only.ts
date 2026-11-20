import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_DIR = path.resolve(SCRIPT_DIR, '..');

const DAILY_REPORT_DIR =
  process.env.DAILY_REPORT_DIR ??
  path.join(PROJECT_DIR, 'reports');

const INPUT_DIR =
  process.env.INPUT_DIR ??
  path.resolve(SCRIPT_DIR, 'input');

export function findTemplateFile() {
  const templatePath = process.env.DAILY_REPORT_TEMPLATE ??
    path.join(PROJECT_DIR, 'templates', 'Tickets_AllZoho.xls');
  if (!fs.existsSync(templatePath)) {
    throw new Error(`Template not found: ${templatePath}. Restore templates/Tickets_AllZoho.xls or set DAILY_REPORT_TEMPLATE.`);
  }
  return templatePath;
}

function clearLine() {
  if (!process.stdout.isTTY) return;
  process.stdout.write('\r\x1b[K');
}

function runPowerShellWithLoading(args: string[]) {
  return new Promise<void>((resolve, reject) => {
    const frames = ['⠋', '⠙', '⠹', '⠸', '⠼', '⠴', '⠦', '⠧', '⠇', '⠏'];
    let frameIndex = 0;
    const startedAt = Date.now();

    const child = spawn('powershell.exe', args, {
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    });

    const timer = setInterval(() => {
      if (!process.stdout.isTTY) return;
      const elapsedSeconds = Math.floor((Date.now() - startedAt) / 1000);
      const frame = frames[frameIndex % frames.length];
      frameIndex++;

      clearLine();
      process.stdout.write(`${frame} Generating daily Zoho report... ${elapsedSeconds}s`);
    }, 120);

    child.stdout.on('data', (data) => {
      clearLine();
      process.stdout.write(data.toString());
    });

    child.stderr.on('data', (data) => {
      clearLine();
      process.stderr.write(data.toString());
    });

    child.on('error', (error) => {
      clearInterval(timer);
      clearLine();
      reject(error);
    });

    child.on('close', (code) => {
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

export async function mergeReport(inputDir = INPUT_DIR, exports?: { today: string; openAll: string }) {
  const ps1Path = path.resolve(SCRIPT_DIR, 'update-zoho-daily-report.ps1');
  const templatePath = findTemplateFile();
  fs.mkdirSync(DAILY_REPORT_DIR, { recursive: true });

  console.log(`Using template: ${templatePath}`);
  console.log(`Using input dir: ${inputDir}`);
  console.log(`Using output dir: ${DAILY_REPORT_DIR}`);
  console.log('');

  await runPowerShellWithLoading([
    '-ExecutionPolicy',
    'Bypass',
    '-File',
    ps1Path,
    '-TemplatePath',
    templatePath,
    '-InputDir',
    inputDir,
    '-OutputDir',
    DAILY_REPORT_DIR,
    ...(exports ? ['-TodayExport', exports.today, '-OpenAllExport', exports.openAll] : []),
  ]);

  console.log('Done.');
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) mergeReport().catch((error) => {
  console.error(error);
  process.exit(1);
});
