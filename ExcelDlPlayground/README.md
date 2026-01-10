# ExcelDlPlayground — Excel-DNA net48 scaffold for DL-in-Excel

This project is a **net48 Excel-DNA** add-in that brings small TorchSharp models into Excel via UDFs. It uses dynamic arrays (spill), async training to keep Excel responsive, and an in-process model registry.

## Environment
- Excel: Microsoft 365 (64-bit)
- Target Framework: net48 (AnyCPU; PlatformTarget x64 in project)
- Excel-DNA: ExcelDna.AddIn 1.9.0
- TorchSharp: TorchSharp-cpu 0.105.2, libtorch-cpu-win-x64 2.7.1.0 (CPU-only)
- Additional references: `System.IO.Compression`, `Microsoft.CSharp`

## What’s implemented
- `DL.MODEL_CREATE(description)` → model_id (default MLP: Linear → Tanh → Linear)
- `DL.TRAIN(model_id, X, y, opts, trigger)` → training summary (async, trigger guard, recalc throttled per epoch/completion)
  - opts: `epochs`, `lr`, `optim=adam|sgd`, `reset=true`
- `DL.LOSS_HISTORY(model_id)` → {epoch, loss}
- `DL.PREDICT(model_id, X)` → logits
- Inspection: `DL.WEIGHTS`, `DL.GRADS`, `DL.ACTIVATIONS`
- Persistence: `DL.SAVE(model_id, path)`, `DL.LOAD(path)` (.dlzip contains model.pt + metadata.json)
- Debug: `DL.TRIGGER_KEY`, `DL.TORCH_TEST`, `DL.TORCH_TEST_DETAIL`, `DL.TORCH_NATIVE_CHECK`, `DL.LOG_PATH`, `DL.LOG_WRITE_TEST`
- Utilities: `WaitAsync(ms)`, `SayHello`, `MatMul`

## Quick Start (Excel test sheet)
1) Data:
```
A2:B4 = X (3x2)
C2:C4 = y (3x1)
```
2) `E2: =DL.MODEL_CREATE("mlp:in=2,hidden=8,out=1")`
3) Trigger cell `Z1` (change to retrain); optional `AA1: =DL.TRIGGER_KEY($Z$1)`
4) Train `E4: =DL.TRAIN(E2, A2:B4, C2:C4, "epochs=20", $Z$1)`
5) Loss history `E8: =DL.LOSS_HISTORY(E2)`
6) Predict `=DL.PREDICT(E2, A2:B4)`
7) Inspect `=DL.WEIGHTS(E2,1)`, `=DL.GRADS(E2,1)`, `=DL.ACTIVATIONS(E2,A2:B4,1)`
8) Save/load `=DL.SAVE(E2, "C:\\Temp\\xor.dlzip")`, `=DL.LOAD("C:\\Temp\\xor.dlzip")`

## TorchSharp setup
- Natives (`torch_cpu.dll`, `LibTorchSharp.dll`, etc.) copied to `bin\Debug\net48` by `CopyTorchNativeBinaries` target.
- `EnsureTorch()` sets PATH/TORCHSHARP_HOME and preloads `LibTorchSharp.dll` then `torch_cpu.dll`.
- State dict save uses reflection-based `torch.save`; dynamic calls require `Microsoft.CSharp`.

## Build / Debug
- Restore: `dotnet restore ExcelDlPlayground/ExcelDlPlayground.csproj -r win-x64`
- Build: `dotnet build ExcelDlPlayground/ExcelDlPlayground.csproj -c Debug`
- Launch (VS Debug): starts EXCEL.EXE with `ExcelDlPlayground-AddIn64.xll`
- Log file: `bin\Debug\net48\ExcelDlPlayground.log` (fallback `%TEMP%`)

## Behaviors / gotchas
- Trigger guard: training skipped if trigger unchanged.
- Recalc: each epoch/completion queues a throttled `xlcCalculateNow`; workbook-wide recalc is intentionally avoided (past crashes).
- Volatile inspectors: STATUS/LOSS_HISTORY/WEIGHTS/GRADS/ACTIVATIONS refresh on recalcs.
- CPU-only; 32-bit Excel not supported.

## Next milestones
- Better errors for malformed ranges/layer references.
- More optimizer/loss options.
- Expanded inspection examples (layer index vs name).

## Design principles
- Range-in / spill-out UDFs (batch-oriented)
- Explicit, triggered training
- No COM calls from background threads
- Simple in-process add-in (no VSTO)
