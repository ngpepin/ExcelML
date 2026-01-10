# AGENTS.md

## Project: ExcelDlPlayground (Excel-DNA + net48)

### Fast start checklist (do first)
1) Close Excel.exe.
2) `dotnet restore ExcelDlPlayground/ExcelDlPlayground.csproj -r win-x64`
3) `dotnet build ExcelDlPlayground/ExcelDlPlayground.csproj -c Debug`
4) Load `bin\Debug\net48\ExcelDlPlayground-AddIn64.xll` in 64-bit Excel (or F5 in VS; Debug settings already point to EXCEL.EXE).
5) Verify natives present in `bin\Debug\net48`: `torch_cpu.dll`, `LibTorchSharp.dll` (copied by `CopyTorchNativeBinaries`).
6) In Excel run `=DL.TORCH_TEST_DETAIL()` → should return “torch ok: …”.
7) Ensure Excel is fully closed before rebuilding (XLL locking will break pack/copy).

### Environment
- Excel: Microsoft 365 64-bit
- Target: net48, PlatformTarget x64, Prefer32Bit=false
- NuGet: ExcelDna.AddIn 1.9.0; TorchSharp-cpu 0.105.2; libtorch-cpu-win-x64 2.7.1.0
- Restore path forced: `%USERPROFILE%\.nuget\packages`
- Assemblies referenced: `System.IO.Compression`, `Microsoft.CSharp`

### TorchSharp gotchas (solved)
- Both `torch_cpu.dll` and `LibTorchSharp.dll` must sit beside `TorchSharp.dll` in output.
- csproj `CopyTorchNativeBinaries` globs natives from:
  - `$(PkgLibtorch_cpu_win_x64)\runtimes\win-x64\native\*.dll`
  - `$(PkgTorchSharp_cpu)\runtimes\win-x64\native\*.dll`
  - `$(PkgTorchSharp)\runtimes\win-x64\native\*.dll`
- `EnsureTorch()` sets PATH + TORCHSHARP_HOME to output folder, preloads `LibTorchSharp.dll` then `torch_cpu.dll` via LoadLibrary.
- TorchSharp calls use `dynamic` for `forward`/`load_state_dict`; **Microsoft.CSharp** reference is required to avoid CS0656.
- Save/load uses reflection-based `torch.save` for state dicts (TorchSharp 0.105 only exposes Tensor overload).

### Trigger/no-op behavior
- `DL.TRAIN` skips when `trigger` unchanged; returns array with `skipped`/`last`/`curr`.
- To retrain: change trigger cell (Z1) and recalc; `AA1: =DL.TRIGGER_KEY($Z$1)` must change.

### Core worksheet repro
- E2: `=DL.MODEL_CREATE("mlp:in=2,hidden=8,out=1")`
- Z1: set 1 (then 2,3,4…)
- AA1: `=DL.TRIGGER_KEY($Z$1)`
- E4: `=DL.TRAIN(E2, A2:B4, C2:C4, "epochs=20", $Z$1)`
- E8: `=DL.LOSS_HISTORY(E2)`
- Optional inference: `=DL.PREDICT(E2, A2:B4)`
- Inspect: `=DL.WEIGHTS(E2,1)` or `=DL.ACTIVATIONS(E2,A2:B4,1)` or `=DL.GRADS(E2,1)` (after training)
- Save/Load: `=DL.SAVE(E2, "C:\\Temp\\xor.dlzip")`, `=DL.LOAD("C:\\Temp\\xor.dlzip")`

### Debug helpers
- Torch: `DL.TORCH_TEST`, `DL.TORCH_TEST_DETAIL`, `DL.TORCH_NATIVE_CHECK` (reports missing native DLLs)
- Logging: `DL.LOG_PATH`, `DL.LOG_WRITE_TEST`; log at `bin\Debug\net48\ExcelDlPlayground.log` (fallback `%TEMP%`).

### Solution/config gotcha
- Solution previously referenced non-existent configs. Ensure solution uses `Debug|Any CPU` / `Release|Any CPU` (remove x64 rows) in Configuration Manager.

### Build/launch notes
- VS Debug: StartProgram set to `EXCEL.EXE` with args `"$(TargetDir)ExcelDlPlayground-AddIn64.xll"`; packing skipped in Debug.
- If ExcelDnaPack fails due to locked XLL, close Excel and rebuild.

### Current State (Important)
- TorchSharp training, predict, save/load, weights/grad/activation inspection implemented for small MLPs.
- Trigger guard works (returns skipped when trigger unchanged).
- Native copy + preload resolves `TypeInitializationException`.
- Dynamic forward requires Microsoft.CSharp; missing reference causes CS0656.
- **Refresh behavior:** during training, each epoch queues a throttled `xlcCalculateNow` via `QueueRecalcOnce` (coalesces if a recalc is already queued). Completion also queues a recalc. STATUS/LOSS_HISTORY/WEIGHTS/GRADS/ACTIVATIONS are volatile and refresh on these per-epoch/completion recalcs. Workbook-wide recalc is intentionally avoided (previously crashed Excel).

### Do NOT try (known pitfalls)
- **Loading 32-bit Excel / x86**: TorchSharp native binaries here are win-x64 only; 32-bit Excel will fail to load natives.
- **Relying on `C:\nuget` global cache**: locked files prevented native downloads; we force `%USERPROFILE%\.nuget\packages` instead.
- **Skipping `EnsureTorch()` preload**: TorchSharp may not find `LibTorchSharp.dll`/`torch_cpu.dll` even if present; preload sets PATH/TORCHSHARP_HOME and LoadLibrary order.
- **Trigger-based training without changing trigger cell**: `DL.TRAIN` will intentionally return `skipped` if trigger unchanged.
- **Using Debug config with x64 solution mappings that don’t exist**: leads to “project configuration does not exist” warnings; use Any CPU configs as documented.
- **Expecting GPU**: this build is CPU-only (libtorch-cpu-win-x64).
- **Forcing workbook-wide recalcs on training completion**: previously caused Excel crash; rely on per-epoch/completion throttled `xlcCalculateNow` instead.

### Notes / Known Past Issues
- Missing libtorch natives caused `TypeInitializationException`; fixed via copy glob + preload.
- CS0656 from dynamic TorchSharp calls is fixed by referencing Microsoft.CSharp.

### Next steps
- Harden error messages for malformed ranges.
- Add docs/examples for layer selection (`index` vs `name`) in inspection UDFs.
- Consider exposing more loss/optimizer options if needed.
