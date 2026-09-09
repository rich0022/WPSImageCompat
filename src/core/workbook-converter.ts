import { WorkbookError } from '../utils/errors';
import { detectWorkbook } from './workbook-detector';
import { renderImages } from './image-renderer';
import type { PreviewSettings } from './image-layout';
import { checkCompatibility } from './compatibility/controller';
import { planFormulaConversions, type FormulaConversionCandidate } from './conversion-planner';
import { parseDispimgFormula } from './dispimg-formula';
import type { DispimgCell } from '../types/wps';

export interface WorkbookConversionOptions extends PreviewSettings {
  convertDispimg: boolean;
  freezeUnsupported: boolean;
  freezeExternal: boolean;
}
export interface ConversionIssue { location: string; code: string; message: string }
export interface WorkbookConversionResult {
  convertedImages: number;
  clearedDispimgFormulas: number;
  frozenFormulas: number;
  unresolvedBrokenReferences: number;
  skipped: number;
  issues: ConversionIssue[];
}
export type ConversionStage = 'reading' | 'checking' | 'convertingImages' | 'convertingFormulas';
const defaults = { detect: detectWorkbook, render: renderImages, compatibility: checkCompatibility };

function issue(location: string, code: string, message: string): ConversionIssue { return { location, code, message }; }
async function clearConvertedDispimg(cells: DispimgCell[]): Promise<{ cleared: number; issues: ConversionIssue[] }> {
  const result = { cleared: 0, issues: [] as ConversionIssue[] };
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const byName = new Map(sheets.items.map(sheet => [sheet.name, sheet]));
    for (const cell of cells) {
      const location = `${cell.worksheetName}!${cell.address}`;
      try {
        const sheet = byName.get(cell.worksheetName);
        if (!sheet) throw new WorkbookError('SHEET_CHANGED', 'Worksheet no longer exists.');
        const range = sheet.getRange(cell.address);
        range.load('formulas,values');
        await context.sync();
        if (range.formulas[0]?.[0] === range.values[0]?.[0] || parseDispimgFormula(range.formulas[0]?.[0]) !== cell.imageId) {
          throw new WorkbookError('CELL_CHANGED', 'DISPIMG formula changed before conversion.');
        }
        range.clear(Excel.ClearApplyTo.contents);
        await context.sync();
        result.cleared++;
      } catch (error) {
        result.issues.push(issue(location, error instanceof WorkbookError ? error.code : 'CONVERSION_WRITE_FAILED',
          error instanceof Error ? error.message : 'Could not remove the converted DISPIMG formula.'));
      }
    }
  });
  return result;
}
async function freezeFormulaValues(candidates: FormulaConversionCandidate[]): Promise<{ frozen: number; issues: ConversionIssue[] }> {
  const result = { frozen: 0, issues: [] as ConversionIssue[] };
  await Excel.run(async context => {
    const sheets = context.workbook.worksheets;
    sheets.load('items/name');
    await context.sync();
    const byName = new Map(sheets.items.map(sheet => [sheet.name, sheet]));
    for (const candidate of candidates) {
      const location = `${candidate.worksheetName}!${candidate.address}`;
      try {
        const sheet = byName.get(candidate.worksheetName);
        if (!sheet) throw new WorkbookError('SHEET_CHANGED', 'Worksheet no longer exists.');
        const range = sheet.getRange(candidate.address);
        range.load('formulas,values,valueTypes');
        await context.sync();
        if (range.formulas[0]?.[0] !== candidate.formula) throw new WorkbookError('CELL_CHANGED', 'Formula changed before conversion.');
        if (range.valueTypes[0]?.[0] === 'Error') throw new WorkbookError('FORMULA_VALUE_ERROR', 'Current formula result is an Excel error and cannot be frozen as a usable value.');
        range.values = [[range.values[0]?.[0] ?? '']];
        await context.sync();
        result.frozen++;
      } catch (error) {
        result.issues.push(issue(location, error instanceof WorkbookError ? error.code : 'CONVERSION_WRITE_FAILED',
          error instanceof Error ? error.message : 'Could not replace the formula with its current value.'));
      }
    }
  });
  return result;
}

/** Explicit conversion only: permanent images and user-selected value freezing in the current workbook. */
export async function convertWorkbook(options: WorkbookConversionOptions, onStage?: (stage: ConversionStage) => void,
  services = defaults): Promise<WorkbookConversionResult> {
  if (!options.convertDispimg && !options.freezeUnsupported && !options.freezeExternal) {
    throw new WorkbookError('NO_CONVERSION_SELECTED', 'Select at least one conversion option.');
  }
  const result: WorkbookConversionResult = { convertedImages: 0, clearedDispimgFormulas: 0, frozenFormulas: 0,
    unresolvedBrokenReferences: 0, skipped: 0, issues: [] };
  if (options.convertDispimg) {
    onStage?.('reading');
    const detection = await services.detect();
    if (!detection.mappings || detection.resourceError) {
      result.issues.push(issue('Workbook', detection.resourceErrorCode ?? 'RESOURCES_UNAVAILABLE',
        detection.resourceError ?? 'WPS image resources could not be read.'));
    } else {
      try {
        onStage?.('convertingImages');
        const rendered = await services.render(detection.mappings, options, 'convert');
        result.convertedImages += rendered.inserted + rendered.existing;
        result.skipped += rendered.skipped;
        result.issues.push(...rendered.issues.map(item => issue(item.location, item.code ?? 'CONVERSION_IMAGE_FAILED', item.message)));
        const cleared = await clearConvertedDispimg(rendered.renderedCells ?? []);
        result.clearedDispimgFormulas += cleared.cleared;
        result.issues.push(...cleared.issues);
      } catch (error) {
        result.skipped += detection.mappings.length;
        result.issues.push(issue('Workbook', error instanceof WorkbookError ? error.code : 'CONVERSION_IMAGE_FAILED',
          error instanceof Error ? error.message : 'Could not convert WPS images.'));
      }
    }
  }
  if (options.freezeUnsupported || options.freezeExternal) {
    onStage?.('checking');
    const report = await services.compatibility();
    const plan = planFormulaConversions(report, options);
    result.unresolvedBrokenReferences = plan.brokenReferenceCount;
    if (plan.blockedByDetailLimit) {
      result.issues.push(issue('Workbook', 'CONVERSION_DETAIL_LIMIT', 'Too many matching formulas to safely convert from the detailed report. Narrow the workbook before retrying.'));
      result.skipped += plan.candidates.length;
    } else if (plan.candidates.length) {
      onStage?.('convertingFormulas');
      const frozen = await freezeFormulaValues(plan.candidates);
      result.frozenFormulas += frozen.frozen;
      result.skipped += frozen.issues.length;
      result.issues.push(...frozen.issues);
    }
  }
  return result;
}
