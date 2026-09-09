import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fixture } from './fixtures/wps-workbook';
const exec = promisify(execFile);
test('file diagnostic reads the real parser, exports aggregate results and never overwrites its input or report', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'wps-diagnostic-test-'));
  try {
    const source = join(directory, 'private-workbook.xlsx');
    const report = join(directory, 'report.json');
    const zip = fixture();
    zip.file('xl/worksheets/sheet1.xml', '<worksheet><sheetData><row r="1"><c r="A1"><f>DISPIMG("ID_A",1)</f></c></row></sheetData></worksheet>');
    const bytes = await zip.generateAsync({ type: 'nodebuffer' });
    await writeFile(source, bytes);
    const run = (destination: string) => exec(process.execPath, ['--import', 'tsx', 'scripts/diagnose-workbook.ts', source, destination]);
    await run(report);
    const output = await readFile(report, 'utf8');
    const parsed = JSON.parse(output);
    assert.equal(parsed.originalUnchanged, true);
    assert.equal(parsed.nativeExcelTested, false);
    assert.deepEqual(parsed.worksheetResults, [{ part: 'xl/worksheets/sheet1.xml', imageCells: 1, found: 1, missing: 0 }]);
    assert.equal(output.includes('ID_A'), false);
    assert.equal(output.includes('private-workbook'), false);
    await assert.rejects(run(source), /must not overwrite/);
    await assert.rejects(run(report), /EEXIST/);
    assert.deepEqual(await readFile(source), bytes);
    assert.equal(await readFile(report, 'utf8'), output);
  } finally { await rm(directory, { recursive: true, force: true }); }
});
