/* =========================
   UTILITIES
========================= */

function dedupeArray(arr) {
  return (arr || []).filter((v, i, a) => a.indexOf(v) === i);
}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function safeBoolean(value, defaultValue) {
  if (typeof value === 'boolean') return value;
  return defaultValue;
}

function overlapsRange(a, b) {
  const aRow1 = a.getRow();
  const aCol1 = a.getColumn();
  const aRow2 = aRow1 + a.getNumRows() - 1;
  const aCol2 = aCol1 + a.getNumColumns() - 1;

  const bRow1 = b.getRow();
  const bCol1 = b.getColumn();
  const bRow2 = bRow1 + b.getNumRows() - 1;
  const bCol2 = bCol1 + b.getNumColumns() - 1;

  const rowsOverlap = aRow1 <= bRow2 && aRow2 >= bRow1;
  const colsOverlap = aCol1 <= bCol2 && aCol2 >= bCol1;

  return rowsOverlap && colsOverlap;
}

function cleanupOverlappingMergedRanges(targetSheet, targetRange) {
  const mergedRanges = targetSheet.getDataRange().getMergedRanges();
  mergedRanges.forEach(mr => {
    if (overlapsRange(mr, targetRange)) {
      try {
        mr.breakApart();
      } catch (e) {}
    }
  });
}

function getSafeTargetRange(targetRange) {
  try {
    const mergedRanges = targetRange.getMergedRanges();
    if (mergedRanges && mergedRanges.length > 0) {
      return mergedRanges[0];
    }
  } catch (e) {}
  return targetRange;
}
