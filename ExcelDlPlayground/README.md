# README.md

## ExcelDlPlayground — Excel-DNA net48 scaffold for DL-in-Excel

This project is a **net48 Excel-DNA** add-in intended to become a learning-focused deep learning playground inside Excel using:
- dynamic arrays (spill)
- structured references / tables (later)
- async training patterns (so Excel stays responsive)
- an in-process model registry

### Environment
- Excel: Microsoft 365 (64-bit), Version 2507 Build 16.0.19029.20136
- Target Framework: net48
- Excel-DNA: ExcelDna.AddIn (intended v1.9.0)

---

## What’s implemented (scaffold)
### Functions
- `DL.MODEL_CREATE(description) -> model_id`
- `DL.TRAIN(model_id, X, y, opts, trigger) -> training summary (spilled)`
- `DL.LOSS_HISTORY(model_id) -> {epoch, loss} (spilled)`
- `DL.TRIGGER_KEY(trigger) -> debug normalized trigger key`
- `WaitAsync(ms) -> timestamp` (async plumbing test)
- Misc helpers (e.g., `SayHello`, `MatMul`) depending on the project state

### Ribbon
A simple “Deep Learning” tab implemented via Excel-DNA Ribbon XML.

---

## Quick Start (Excel test sheet)

### 1) Enter sample data
Create this table in Excel:

| A | B | C |
|---|---|---|
| X1 | X2 | y |
| 1 | 2 | 0 |
| 2 | 3 | 1 |
| 3 | 4 | 0 |

- X range = `A2:B4`  (3 rows x 2 features)
- y range = `C2:C4`  (3 rows x 1 target)

### 2) Create a model
In `E2`:

```excel
=DL.MODEL_CREATE("mlp:in=2,hidden=8,out=1")
```

This returns a `model_id` string.

### 3) Create a trigger cell
In `Z1`, type:

```excel
1
```

Change it to `2`, `3`, etc. to request retraining.

### 4) Debug trigger key (recommended)

In `AA1`:

```excel
=DL.TRIGGER_KEY($Z$1)
```

This should reflect Z1’s value immediately.

### 5) Train

In `E4`:

```excel
=DL.TRAIN(E2, A2:B4, C2:C4, "epochs=20", $Z$1)
```

Expected behavior:

* First run with a new trigger: returns a summary like `{status, done; epochs, 20; final_loss, ...}`
* Recalc without changing Z1: returns `no-op / trigger unchanged`
* Change Z1: should retrain

### 6) View loss history

In `E8`:

```excel
=DL.LOSS_HISTORY(E2)
```

Spills:

```
epoch | loss
1     | 0.92
2     | 0.8464
...
```

---

## TorchSharp setup (current)

- Packages: `TorchSharp-cpu` 0.105.2, `libtorch-cpu-win-x64` 2.7.1.0
- Natives (`torch_cpu.dll`, `LibTorchSharp.dll`, etc.) are copied from `%USERPROFILE%\.nuget\packages` into `bin\Debug\net48` via the csproj target `CopyTorchNativeBinaries`.
- `EnsureTorch()` sets `PATH`/`TORCHSHARP_HOME` and preloads `LibTorchSharp.dll` then `torch_cpu.dll`.
- Debug UDFs: `DL.TORCH_TEST`, `DL.TORCH_TEST_DETAIL` show init status.

---

## Build / Debug tips

- Use **64-bit Excel**.
- If TorchSharp complains about missing libtorch: confirm `torch_cpu.dll` and `LibTorchSharp.dll` are in `bin\Debug\net48`; re-run `dotnet restore -r win-x64` and rebuild.
- Add-in launch: VS Debug starts `EXCEL.EXE` with `ExcelDlPlayground-AddIn64.xll`.

### Logging (recommended)

If added, check `%TEMP%\ExcelDlPlayground.log` for call traces.

---

## Current Known Issue
Trigger guard now works; TorchSharp init fixed via native copy/preload. If you see `TypeInitializationException`, check native DLLs as above.

---

## Next milestones (after trigger fix)

1. Replace placeholder “fake training” with TorchSharp CPU MLP (e.g., XOR).
2. Add `DL.PREDICT(model_id, X)` for batched inference (spilled).
3. Add inspection functions for learning:
   * `DL.WEIGHTS(model_id, layer)`
   * `DL.ACTIVATIONS(model_id, X, layer)`
   * `DL.GRADS(model_id, layer)`
4. Add a model persistence option (save/load).

---

## Design principles

* Prefer **range-in / spill-out** UDFs (batch operations, not cell-by-cell).
* Training should be **explicit** and **triggered**, never accidental.
* Avoid calling Excel COM from background threads.
* Keep add-in single-process/simple (no VSTO).