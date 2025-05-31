# Notes for Next AI Developer/Assistant

## Project Overview
- **Project Name:** VBA Schedule Aggregator (工程表データ集約マクロ)
- **Current Main Goal:** Stabilize configuration reading (`M02_ConfigReader`) and then proceed to implement/verify core data extraction logic (`M06_DataExtractor`).

## Development History & Key Concepts
- **Configuration Reading (`M02_ConfigReader`):**
  - This module has undergone significant refactoring to handle reading various data types (especially arrays of strings) from Excel ranges robustly.
  - Key Pattern:
    - `ReadRangeToArray(ws, rangeAddress, ...)`: Simplified function that returns a raw `Variant` from the range (or `Empty` on direct read error).
    - `ConvertRawVariantToStringArray(rawData, ..., currentConfig)`: Helper function that intelligently converts the `Variant` from `ReadRangeToArray` into a standardized 1D `String()` array (empty arrays are `(1 To 0)`). This function contains detailed logic for handling scalars, 1D arrays, and 2D vertical arrays (N rows x 1 col). It's crucial for reading lists from the Config sheet. It now includes detailed internal logging controlled by `currentConfig.DebugDetailLevel2Enabled`.
  - Error Handling: Each `Load...` sub-procedure in `M02_ConfigReader` (e.g., `LoadGeneralSettings`, `LoadFilterConditions`) has its own error handler that logs context-specific information (e.g., which `currentItem` was being processed) and then calls `M02Reader_LogAndSetError`. This module-level helper sets the `m_errorOccurred` flag and logs the error, allowing `LoadConfiguration` to gracefully stop further processing and return `False`.
- **Logging Framework:**
  - **Error Log (O7, O45):** Records errors. Output to sheet controlled by `EnableErrorLogSheetOutput`. Falls back to `Debug.Print` if sheet not available.
  - **Filter Log (O6, O44):** Records *only* D-Section filter settings that are actively set. Output to sheet controlled by `EnableSearchConditionLogSheetOutput`. See `M04_LogWriter.WriteFilterLog` and `WriteFilterLogArrayEntryIfSet`.
  - **Operation Log (O5, O42):** Records general macro operations, settings snapshot (excluding filters), and can be used for event/progress logging. Output to sheet controlled by `EnableSheetLogging`. See `M04_LogWriter.WriteOperationLog`.
- **Debug Detail Levels (O4, O8, O9):**
  - `DebugDetailLevel1Enabled` (O4, default TRUE): For critical path debug messages, error tracing. (e.g., "DEBUG_POINT" logs in `MainControl`).
  - `DebugDetailLevel2Enabled` (O8, default FALSE): For detailed variable states, array contents during config loading. (e.g., `DebugPrintArrayState` outputs, F-Section item details in `LoadConfiguration`, internal logs in `ConvertRawVariantToStringArray`).
  - `DebugDetailLevel3Enabled` (O9, default FALSE): For verbose config dumps after successful loading. (e.g., `--- Loaded Configuration Settings ---` block in `LoadConfiguration`).
  These flags control `Debug.Print` and some `WriteErrorLog` calls (using "DEBUG_L2", "DEBUG_ARRAY_STATE" levels).

## Current Status & Potential Issues
- **Persistent Error (as of user feedback prior to subtask group starting at step 20):** "実行時エラー 9: インデックスが有効範囲にありません。" occurring during `M02_ConfigReader.LoadConfiguration`, logged as originating from `LoadProcessPatternDefinition` when processing "Bunrui1List (L129:L167)".
- **Hypothesis & Actions Taken:** The primary hypothesis was that `ReadRangeToArray` and subsequent processing into string arrays was not robust enough for all edge cases (empty ranges, single cell ranges, ranges with all empty cells, unexpected array dimensions from `.value`).
  - `ReadRangeToArray` was simplified to return a raw `Variant`.
  - `ConvertRawVariantToStringArray` was introduced to handle the complex logic of interpreting this `Variant` into a clean `String()`, including detailed `DebugDetailLevel2Enabled` logging of its internal steps.
  - This new pattern was systematically applied to *all* `ReadRangeToArray` call sites in `M02_ConfigReader.bas` for loading list-type settings.
  - Error handling in all `Load...` sub-procedures was made consistent, logging specific items and then propagating errors to `LoadConfiguration`.
  - Debug logging levels were introduced for more targeted analysis.
  - The F-Section debug print logic in `LoadConfiguration` was corrected.
- **Current Expectation:** The "Index out of bounds" error for "Bunrui1List" (and similar errors for other lists) should now be resolved or, if it persists, the new `DebugDetailLevel2Enabled` logging within `ConvertRawVariantToStringArray` should provide very specific information about why the conversion is failing for that particular range/item.

## Next Steps & Recommendations for AI
1.  **Confirm Resolution of "Bunrui1List" Error (Top Priority):**
    - Instruct the user to set Config!O8 (`DebugDetailLevel2Enabled`) to TRUE (and O4, O7 to TRUE).
    - Run the macro.
    - **If the error persists:** Carefully analyze the detailed ErrorLog output, focusing on logs from `ConvertRawVariantToStringArray` for "Bunrui1List". The logs show `TypeName(rawData)`, detected bounds, loop iterations, etc. This data is essential to pinpoint the exact line or condition within `ConvertRawVariantToStringArray` still causing issues. Refine `ConvertRawVariantToStringArray` further based on this.
    - **If the error is resolved:** Proceed to verify other array reads.
2.  **Verify All Config Array Reads:** Systematically check (or ask user to check with `DebugDetailLevel2Enabled = TRUE`) that all other list-based settings in `M02_ConfigReader` load correctly. This includes:
    - `TargetSheetNames` (in `LoadScheduleFileSettings`)
    - `ProcessKeys`, `Kankatsu1List`-`Bunrui3List` (in `LoadProcessPatternDefinition`)
    - All lists in `LoadFilterConditions`
    - `TargetFileFolderPaths`, `FilePatternIdentifiers` (in `LoadTargetFileDefinition`)
    - `HideSheetNames` (in `LoadConfiguration`'s G-Section)
3.  **Robust Input Data Handling (`M06_DataExtractor`):**
    - After `M02_ConfigReader` is stable, shift focus to `M06_DataExtractor`.
    - Implement robust error handling for cases where input schedule files are missing expected sheets, or where cells for Year, Month, Day, or other critical data points are empty or contain invalid data.
    - The macro should log such issues and gracefully skip the problematic file/sheet/data point, rather than halting.
4.  **FilterLog Content Review (`M04_LogWriter.WriteFilterLog`):**
    - Confirm with user or by testing that the "Search Conditions Log" (controlled by O6) now correctly logs only filter conditions that have been actively set/modified by the user, not default or empty values. The current implementation logs non-default `WorkerFilterLogic`, non-empty string filters, lists if they contain any values (via `WriteFilterLogArrayEntryIfSet`), and `NinzuFilter` if it wasn't originally empty.
5.  **Continuous Documentation Update:**
    - Ensure `docs/04_Config_Sheet_Definition.md` is accurate for all settings, especially debug/logging flags.
    - Update these `NOTES_FOR_NEXT_AI.md` with new status or resolved issues.

## Recommended Debug Settings (Config Sheet)
- **For diagnosing potential remaining array conversion issues:**
  - O4 (`DebugDetailLevel1Enabled`): TRUE
  - O7 (`EnableErrorLogSheetOutput`): TRUE
  - O8 (`DebugDetailLevel2Enabled`): TRUE (Crucial for `ConvertRawVariantToStringArray` internal logs)
  - O9 (`DebugDetailLevel3Enabled`): FALSE or TRUE (L3 gives full config dump on success, which can be useful for overall verification).
- **For general development/testing once current error is fixed:**
  - O4: TRUE
  - O5, O6, O7: TRUE (to ensure all log sheets are active and capturing info)
  - O8: TRUE or FALSE depending on need for detailed variable tracing.
  - O9: FALSE.

Good luck!
```
