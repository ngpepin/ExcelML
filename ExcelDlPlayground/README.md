# README.md

## ExcelDlPlayground — Excel-DNA net48 scaffold for DL-in-Excel

This project is a **net48 Excel-DNA** add-in intended to become a learning-focused deep learning playground inside Excel using:
- dynamic arrays (spill)
- structured references / tables (later)
- async training patterns (so Excel stays responsive)
- an in-process model registry

### Environment
- Excel: Microsoft 365 (64-bit)
- Target Framework: net48 (AnyCPU platform, PlatformTarget x64 in project)
- Excel-DNA: ExcelDna.AddIn 1.9.0
- TorchSharp packages: TorchSharp-cpu 0.105.2, libtorch-cpu-win-x64 2.7.1.0
- Additional references: `System.IO.Compression`, `Microsoft.CSharp` (needed for dynamic TorchSharp calls)

---

## What’s implemented
### Functions
- `DL.MODEL_CREATE(description) -> model_id`
- `DL.TRAIN(model_id, X, y, opts, trigger) -> training summary (spilled)` with trigger guard (no-op if trigger unchanged)
  - opts supports `epochs`, `lr`, `optim` (`adam`/`sgd`), and `reset=true` to reinitialize weights/optimizer before training
- `DL.LOSS_HISTORY(model_id) -> {epoch, loss} (spilled)`
- `DL.PREDICT(model_id, X) -> outputs (spilled)`
- Inspection: `DL.WEIGHTS(model_id, layer)`, `DL.GRADS(model_id, layer)`, `DL.ACTIVATIONS(model_id, X, layer)`
- Persistence: `DL.SAVE(model_id, path)`, `DL.LOAD(path) -> model_id`
- Debug: `DL.TRIGGER_KEY(trigger)`, `DL.TORCH_TEST`, `DL.TORCH_TEST_DETAIL`, `DL.TORCH_NATIVE_CHECK`, `DL.LOG_PATH`, `DL.LOG_WRITE_TEST`
- Utilities: `WaitAsync(ms)`, `SayHello`, `MatMul`

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
In `Z1`, type a value (e.g., `1`). Change it to `2`, `3`, etc. to retrain.

### 4) Debug trigger key (recommended)

In `AA1`:

```excel
=DL.TRIGGER_KEY($Z$1)
```

This should reflect Z1’s value immediately.

### 5) Train

In `E4`:

```excel
=DL.TRAIN(E2, A2:B4, C2:C4, "epochs=20,reset=true", $Z$1)
```

Expected behavior:

* First run with a new trigger: returns a summary like `{status, done; epochs, 20; final_loss, ...}`
* Recalc without changing Z1: returns `skipped`/`trigger unchanged`
* Change Z1: retrains (add `reset=true` if you want a fresh weight init each time)

### 6) View loss history

In `E8`:

```excel
=DL.LOSS_HISTORY(E2)
```

Spills a 2-column table of epoch/loss.

### 7) Predict (optional)

```excel
=DL.PREDICT(E2, A2:B4)
```

Spills logits for each row of X.

### 8) Inspect layers (optional)

- Weights: `=DL.WEIGHTS(E2,1)` (or layer name)
- Gradients: `=DL.GRADS(E2,1)` (after training)
- Activations: `=DL.ACTIVATIONS(E2, A2:B4, 1)`

### 9) Save and reload a model

Save after training:

```excel
=DL.SAVE(E2, "C:\Temp\xor-model.dlzip")
```

Reload later (returns the stored `model_id`):

```excel
=DL.LOAD("C:\Temp\xor-model.dlzip")
```

---

## TorchSharp setup (current)

- Natives (`torch_cpu.dll`, `LibTorchSharp.dll`, etc.) are copied into `bin\Debug\net48` via the csproj target `CopyTorchNativeBinaries` (libtorch-cpu-win-x64, torchsharp-cpu, torchsharp runtimes).
- `EnsureTorch()` sets `PATH`/`TORCHSHARP_HOME` and preloads `LibTorchSharp.dll` then `torch_cpu.dll`.
- State dict save uses reflection-based `torch.save` (TorchSharp 0.105 only exposes Tensor overloads).
- Dynamic `forward`/`load_state_dict` calls require the `Microsoft.CSharp` reference (to avoid CS0656).

---

## Build / Debug tips

- Use **64-bit Excel**.
- Restore/build: `dotnet restore ExcelDlPlayground/ExcelDlPlayground.csproj -r win-x64` then `dotnet build ... -c Debug`.
- If TorchSharp complains about missing libtorch: confirm `torch_cpu.dll` and `LibTorchSharp.dll` are in `bin\Debug\net48`; re-run restore/build with `-r win-x64` and rebuild.
- Add-in launch: VS Debug starts `EXCEL.EXE` with `ExcelDlPlayground-AddIn64.xll`.

### Logging

Check `bin\Debug\net48\ExcelDlPlayground.log` (fallback `%TEMP%`).

---

## Known behaviors / issues

- Trigger guard prevents accidental retrain; change the trigger cell to retrain.
- Build requires `Microsoft.CSharp` for dynamic TorchSharp calls; missing it causes CS0656.
- CPU-only build (libtorch-cpu-win-x64); x86 Excel not supported.

---

## Next milestones

- Improve error reporting for malformed ranges or layer references.
- Add richer optimizer/loss options if needed.
- Expand examples for inspection UDFs (layer index vs name).

---

## Design principles

* Prefer **range-in / spill-out** UDFs (batch operations, not cell-by-cell).
* Training should be **explicit** and **triggered**, never accidental.
* Avoid calling Excel COM from background threads.
* Keep add-in single-process/simple (no VSTO).
