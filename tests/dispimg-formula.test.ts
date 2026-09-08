import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseDispimgFormula, cellAddress } from '../src/core/dispimg-formula';

test('recognizes WPS and Excel compatibility formulas', () => {
  for (const formula of ['=DISPIMG("ID_123",1)', '=@_xlfn.DISPIMG("ID_123",1)',
    '=_xlfn.DISPIMG("ID_123",1)', '=@DISPIMG("ID_123",1)',
    ' = dispimg ( "ID_123" ; 1 ) ']) {
    assert.equal(parseDispimgFormula(formula), 'ID_123');
  }
});
test('does not detect text, nested functions, references or incomplete formulas', () => {
  for (const value of [null, 2, 'DISPIMG("ID_1",1)', '\'=DISPIMG("ID_1",1)',
    '=IF(TRUE,DISPIMG("ID_1",1),0)', '=DISPIMG(A1,1)', '=DISPIMG("",1)',
    '=DISPIMG("ID_1",1)+2', '=DISPIMG("ID_1",', '=OTHERDISPIMG("ID_1",1)']) {
    assert.equal(parseDispimgFormula(value), undefined);
  }
});
test('decodes escaped quotes without changing case of IDs', () => {
  assert.equal(parseDispimgFormula('=DISPIMG("ID_a""B",1)'), 'ID_a"B');
});
test('maps zero-based coordinates including worksheet boundaries', () => {
  assert.equal(cellAddress(0, 0), 'A1');
  assert.equal(cellAddress(12, 26), 'AA13');
  assert.equal(cellAddress(1048575, 16383), 'XFD1048576');
  assert.throws(() => cellAddress(-1, 0), RangeError);
});
