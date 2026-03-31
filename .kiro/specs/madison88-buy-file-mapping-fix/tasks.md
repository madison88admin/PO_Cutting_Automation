# Tasks: Madison 88 Buy File Mapping Fix

- [x] 1. Add `Style#` alias for the `product` field
  - Already present in `getFallbackColumnAliases()` — no change needed.

- [x] 2. Implement `product`-field fallback for the PLM style lookup
  - [x] 2.1 When `jdeStyle` is empty AND `product` is non-empty, use `product` as the PLM lookup key
  - [x] 2.2 Only emit "JDE Style missing" when BOTH `jdeStyle` AND `product` are empty
  - [x] 2.3 Do not change behaviour when `jdeStyle` is present

- [x] 3. Fix size column false-positive inference for Madison 88
  - [x] 3.1 Add `'production surcharge confirmation status'` (and similar status-like headers) to `shouldSilentlyIgnoreHeader` so they don't trigger `looksLikeSizeHeader`
  - [x] 3.2 Confirm that when no real size column exists, `useDefaultSizeBucket` kicks in and "One Size" is applied

- [x] 4. Regression check
  - [x] 4.1 Files with explicit `JDE Style` column continue to use `jdeStyle` as primary key
  - [x] 4.2 Files with `jdeStyle` present but empty for a row still emit "JDE Style missing" for that row
