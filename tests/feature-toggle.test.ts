import { test } from 'node:test';
import assert from 'node:assert/strict';
import { featureEnabledPreference } from '../src/taskpane/feature-toggle';

test('feature toggle defaults to enabled and only a persisted false pauses it', () => {
  assert.equal(featureEnabledPreference(null), true);
  assert.equal(featureEnabledPreference(undefined), true);
  assert.equal(featureEnabledPreference('true'), true);
  assert.equal(featureEnabledPreference('false'), false);
});
