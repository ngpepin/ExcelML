# AGENTS.md

## Project: ExcelDlPlayground (Excel-DNA + net48)

### Fast start checklist (do first)
1) Close Excel.exe.
2) `dotnet restore ExcelDlPlayground/ExcelDlPlayground.csproj -r win-x64`
3) `dotnet build ExcelDlPlayground/ExcelDlPlayground.csproj -c Debug`
4) Load `bin\Debug\net48\ExcelDlPlayground-AddIn64.xll` in 64-bit Excel (or F5 in VS; Debug settings already point to EXCEL.EXE).
5) Verify natives in `bin\Debug\net48`: `torch_cpu.dll`, `LibTorchSharp.dll` (copied by `CopyTorchNativeBinaries`).
6) In Excel run `=DL.TORCH_TEST_DETAIL()` → should return “torch ok: …”.
7) Ensure Excel is fully closed before rebuilding (XLL locking will break pack/copy).

### Environment
- Excel: Microsoft 365 64-bit
- Target: net48, PlatformTarget x64, Prefer32Bit=false
- NuGet: ExcelDna.AddIn 1.9.0; TorchSharp-cpu 0.105.2; libtorch-cpu-win-x64 2.7.1.0
- Restore path forced: `%USERPROFILE%\.nuget\packages`
- Assemblies referenced: `System.IO.Compression`, `Microsoft.CSharp`

### TorchSharp gotchas (solved)
- `torch_cpu.dll` and `LibTorchSharp.dll` must sit beside `TorchSharp.dll` in output.
- `CopyTorchNativeBinaries` globs natives from libtorch-cpu-win-x64, torchsharp-cpu, torchsharp runtimes.
- `EnsureTorch()` sets PATH + TORCHSHARP_HOME and preloads `LibTorchSharp.dll` then `torch_cpu.dll` via LoadLibrary.
- Dynamic TorchSharp calls use `dynamic`; **Microsoft.CSharp** reference required to avoid CS0656.
- Save/load uses reflection-based `torch.save` for state dicts (Tensor overload only exposed in 0.105).

### Trigger/no-op behavior
- `DL.TRAIN` skips when `trigger` unchanged; returns `skipped`/`last`/`curr`.
- To retrain: change trigger cell (e.g., Z1) and recalc; `AA1: =DL.TRIGGER_KEY($Z$1)` must change.

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
- Torch: `DL.TORCH_TEST`, `DL.TORCH_TEST_DETAIL`, `DL.TORCH_NATIVE_CHECK`
- Logging: `DL.LOG_PATH`, `DL.LOG_WRITE_TEST`; log at `bin\Debug\net48\ExcelDlPlayground.log` (fallback `%TEMP%`).

### Solution/config gotcha
- Ensure solution configs use `Debug|Any CPU` / `Release|Any CPU` (remove stray x64 mappings).

### Build/launch notes
- VS Debug: StartProgram = `EXCEL.EXE` with args `"$(TargetDir)ExcelDlPlayground-AddIn64.xll"`; packing skipped in Debug.
- If ExcelDnaPack fails due to locked XLL, close Excel and rebuild.

### Current State (Important)
- TorchSharp training, predict, save/load, weights/grad/activation inspection implemented for small MLPs.
- Trigger guard works (returns skipped when trigger unchanged).
- Native copy + preload resolves `TypeInitializationException`.
- Dynamic forward requires Microsoft.CSharp; missing reference causes CS0656.
- **Refresh behavior:** each epoch queues a throttled `xlcCalculateNow` via `QueueRecalcOnce`; completion also queues one. Workbook-wide recalc is avoided (previously crashed Excel). STATUS/LOSS_HISTORY/WEIGHTS/GRADS/ACTIVATIONS are volatile and refresh on these recalcs.

### Do NOT try (known pitfalls)
- Loading 32-bit Excel / x86 (natives are win-x64 only).
- Relying on `C:\nuget` global cache (locked files); use `%USERPROFILE%\.nuget\packages`.
- Skipping `EnsureTorch()` preload (PATH/TORCHSHARP_HOME + LoadLibrary ordering matters).
- Trigger-based training without changing trigger cell (TRAIN will return `skipped`).
- Using Debug configs with non-existent x64 mappings (fix in Configuration Manager).
- Expecting GPU (CPU-only build).
- Forcing workbook-wide recalc; rely on throttled `xlcCalculateNow`.

### Notes / Known Past Issues
- Missing libtorch natives caused `TypeInitializationException`; fixed via copy + preload.
- CS0656 from dynamic TorchSharp calls fixed by Microsoft.CSharp reference.

### Next steps
- Harden error messages for malformed ranges.
- Add docs/examples for layer selection (`index` vs `name`) in inspection UDFs.
- Consider exposing more loss/optimizer options if needed.
