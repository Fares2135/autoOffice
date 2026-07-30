import { describe, it, expect } from 'vitest';
import { unappliedColumns, hasProps, PT_PER_INCH } from './edit.ts';

describe('PT_PER_INCH', () => {
  it('turns the widths users ask for into the units Word wants', () => {
    expect(0.6 * PT_PER_INCH).toBeCloseTo(43.2);
  });
});

describe('unappliedColumns', () => {
  const target = 43.2; // 0.6"

  it('is empty when Word kept every width', () => {
    expect(unappliedColumns([86.4, 43.2, 43.2, 43.2], [1, 2, 3], target)).toEqual([]);
  });

  it('tolerates Word rounding the value slightly', () => {
    expect(unappliedColumns([43.25, 43.15], [0, 1], target)).toEqual([]);
  });

  // The silent failure this exists for: an auto-fit table takes the assignment,
  // reports success, and snaps back.
  it('names the columns that snapped back', () => {
    expect(unappliedColumns([86.4, 86.4, 43.2], [0, 1, 2], target)).toEqual([0, 1]);
  });

  it('counts a missing column as unapplied rather than skipping it', () => {
    expect(unappliedColumns([43.2], [0, 1], target)).toEqual([1]);
  });

  it('claims nothing when the widths could not be read back', () => {
    expect(unappliedColumns(null, [0, 1], target)).toEqual([]);
  });
});

describe('hasProps', () => {
  it('is false for an empty request', () => {
    expect(hasProps({})).toBe(false);
  });

  it('is false when every property is undefined', () => {
    expect(hasProps({ bold: undefined, size: undefined })).toBe(false);
  });

  it('is true for a value that is falsy but meant, like bold: false', () => {
    expect(hasProps({ bold: false })).toBe(true);
  });

  it('is true for a normal request', () => {
    expect(hasProps({ color: '#C00000' })).toBe(true);
  });
});
