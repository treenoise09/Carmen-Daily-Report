import { chromium } from '@playwright/test';
import path from 'node:path';
const context = await chromium.launchPersistentContext(path.resolve('playwright/.auth/zoho-profile'), { headless: false, viewport: { width: 1600, height: 900 } });
try {
 const page = context.pages()[0] ?? await context.newPage();
 await page.goto('https://desk.zoho.com/agent/carmensoftware/carmen-software-support/reports/details/483929000029190842');
 await page.getByRole('button', {name: 'Create Chart', exact:true}).waitFor({timeout:60000});
 const later = page.getByText('Not Now', {exact:true});
 if(await later.isVisible()) await later.click();
 console.log(await page.getByRole('main').getByRole('button').evaluateAll(nodes => nodes.map(n=>n.outerHTML)));
} finally { await context.close(); }
