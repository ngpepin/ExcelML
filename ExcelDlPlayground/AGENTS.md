# AGENTS.md

## Project: ExcelDlPlayground (Excel-DNA + net48)

### Fast start checklist (do first)
1) Close Excel.exe.
2) `dotnet restore ExcelDlPlayground/ExcelDlPlayground.csproj -r win-x64`
3) `dotnet build ExcelDlPlayground/ExcelDlPlayground.csproj -c Debug`
4) Load `bin\Debug\net48\ExcelDlPlayground-AddIn64.xll` in 64-bit Excel (or F5 in VS; Debug settings already point to EXCEL.EXE).
5) Verify natives present in `bin\Debug\net48`: `torch_cpu.dll`, `LibTorchSharp.dll`.
6) In Excel run `=DL.TORCH_TEST_DETAIL()` → should return “torch ok: …”.
7) Ensure Excel is fully closed before rebuilding (XLL locking will break pack/copy).

### Environment
- Excel: Microsoft 365 64-bit
- Target: net48, PlatformTarget x64, Prefer32Bit=false
- NuGet: ExcelDna.AddIn 1.9.0; TorchSharp-cpu 0.105.2; libtorch-cpu-win-x64 2.7.1.0
- Restore path forced: `%USERPROFILE%\.nuget\packages`

### TorchSharp gotchas (solved)
- Needed both `torch_cpu.dll` and `LibTorchSharp.dll` beside `TorchSharp.dll`.
- csproj `CopyTorchNativeBinaries` globs natives from:
  - `$(PkgLibtorch_cpu_win_x64)\runtimes\win-x64\native\*.dll`
  - `$(PkgTorchSharp_cpu)\runtimes\win-x64\native\*.dll`
  - `$(PkgTorchSharp)\runtimes\win-x64\native\*.dll`
- `EnsureTorch()` sets PATH + TORCHSHARP_HOME to output folder, preloads `LibTorchSharp.dll` then `torch_cpu.dll` via LoadLibrary.
- If `TypeInitializationException` reappears: confirm those DLLs exist in output, rerun restore/build with `-r win-x64`, reload XLL.
- If upgrading TorchSharp, verify native package versions explicitly — transitive resolution is unreliable under net48 + Excel.

### Trigger/no-op behavior
- `DL.TRAIN` skips when `trigger` unchanged; returns array with `skipped`/`last`/`curr`.
- To retrain: change trigger cell (Z1) and recalc; `AA1: =DL.TRIGGER_KEY($Z$1)` must change.

### Core worksheet repro
- E2: `=DL.MODEL_CREATE("mlp:in=2,hidden=8,out=1")`
- Z1: set 1 (then 2,3,4…)
- AA1: `=DL.TRIGGER_KEY($Z$1)`
- E4: `=DL.TRAIN(E2, A2:B4, C2:C4, "epochs=20", $Z$1)`
- E8: `=DL.LOSS_HISTORY(E2)`

### Debug helpers
- Torch: `DL.TORCH_TEST`, `DL.TORCH_TEST_DETAIL`, `DL.TORCH_NATIVE_CHECK` (reports missing native DLLs)
- Logging: `DL.LOG_PATH`, `DL.LOG_WRITE_TEST`; log at `bin\Debug\net48\ExcelDlPlayground.log` (fallback `%TEMP%`).

### Solution/config gotcha
- Solution previously referenced non-existent configs. Ensure solution uses `Debug|Any CPU` / `Release|Any CPU` (remove x64 rows) in Configuration Manager.

### Build/launch notes
- VS Debug: StartProgram set to `EXCEL.EXE` with args `"$(TargetDir)ExcelDlPlayground-AddIn64.xll"`; packing skipped in Debug.
- If ExcelDnaPack fails due to locked XLL, close Excel and rebuild.

### Next steps (future work)
- Swap fake training loop with real TorchSharp MLP training.
- Add `DL.PREDICT`, inspection UDFs (weights/activations/grads), persistence.

### Goal
Build an Excel-DNA add-in that supports an Excel-first deep learning learning environment:
- In-process **model registry** with `DL.MODEL_CREATE` returning a `model_id` string
- Trigger-token training to avoid accidental retrains: `DL.TRAIN(model_id, X, y, opts, trigger)`
- Spilled-array loss history: `DL.LOSS_HISTORY(model_id)`
- Later: replace fake training loop with TorchSharp (CPU) and add inference.

### Current State (Important)
The add-in loads and ribbon works. Dynamic arrays spill. Model registry works. Loss history spills.
Trigger guard works (returns skipped when trigger unchanged).
TorchSharp now initializes successfully after adding native copy + preload (see Lessons Learned).

### Required Excel behaviors
1. `DL.MODEL_CREATE(...)` returns a stable model ID.
2. `DL.TRAIN(..., trigger)` runs only when `trigger` changes.
3. `DL.LOSS_HISTORY(model_id)` spills epoch/loss table and updates after training.
4. Changing trigger cell (e.g., Z1: 1→2→3) retrains reliably.

### Notes / Known Past Issues
- Missing libtorch natives caused `TypeInitializationException`; fixed via copy glob + preload.
- Solution config warning: align solution to Debug/Release AnyCPU (remove x64 rows) if it reappears.

### Do NOT try (known pitfalls)
- **Loading 32-bit Excel / x86**: TorchSharp native binaries here are win-x64 only; 32-bit Excel will fail to load natives.
- **Using libtorch-cpu-win-x64 2.7.1.0 without copying natives**: TorchSharp throws `TypeInitializationException` because `torch_cpu.dll`/`LibTorchSharp.dll` are missing next to `TorchSharp.dll`.
- **Relying on `C:\nuget` global cache**: locked files prevented native downloads; we force `%USERPROFILE%\.nuget\packages` instead.
- **Skipping `EnsureTorch()` preload**: TorchSharp may not find `LibTorchSharp.dll`/`torch_cpu.dll` even if present; preload sets PATH/TORCHSHARP_HOME and LoadLibrary order.
- **Trigger-based training without changing trigger cell**: `DL.TRAIN` will intentionally return `skipped` if trigger unchanged.
- **Using Debug config with x64 solution mappings that don’t exist**: leads to “project configuration does not exist” warnings; use Any CPU configs as documented.
