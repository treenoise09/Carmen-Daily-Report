import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(new URL('./audit-preview.xlsx', import.meta.url).pathname.replace(/^\//, '')));
console.log((await wb.inspect({kind:'region',sheetId:'Open_All',range:'A15:K30',maxChars:2000,tableMaxRows:16,tableMaxCols:11})).ndjson);
const preview = await wb.render({sheetName:'Open_All',range:'A15:K30',scale:1.5,format:'png'});
await fs.writeFile(new URL('./audit-preview.png',import.meta.url),new Uint8Array(await preview.arrayBuffer()));
