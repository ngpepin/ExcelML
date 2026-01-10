Project: **ExcelDlPlayground**
------------------------------

**Excel-native deep learning with Excel-DNA + TorchSharp (net48)**

This project embeds **stateful machine learning models** inside Excel while respecting Excel’s recalculation engine, threading model, and UI constraints.

> ⚠️ This is **not** a normal UDF project.  
> Correct behavior depends on following a small set of **non-obvious but proven patterns** documented below.

* * *

1. Core Design Philosophy (Read First)
---------------------------------------

### Excel is the orchestrator, not the runtime

Excel:

* Recalculates aggressively
* Replays formulas on open
* Has **no concept of state** between calls
* Runs UDFs on the **UI thread unless async/observable**

Therefore:

* **All persistent state lives outside Excel formulas**
* Excel formulas are _signals_, not commands
* Training, prediction, and inspection must be **idempotent**

* * *

2. Three Proven UDF Categories (Critical)
-----------------------------------------

### ① **Creation / Identity (Idempotent, Cached)**

Examples:

* `DL.MODEL_CREATE`
* `DL.SESSION_ID`

**Rules**

* Must return the _same value_ when inputs are unchanged
* Must not recreate objects on every recalc
* Must be safe on workbook open

**Pattern**

* Caller-scoped cache (keyed by caller cell + trigger)
* Optional trigger to force regeneration
* Session-stable identity when needed

### ② **Triggered Work (Async, Fire-Once)**

Examples:

* `DL.TRAIN`

**Rules**

* Must not run twice for the same trigger
* Must not block Excel UI
* Must tolerate recalculation storms

**Pattern**

* Trigger normalization (`TriggerKey`)
* LastTriggerKey guard
* `ExcelAsyncUtil.RunTask`
* Non-blocking `SemaphoreSlim`

### ③ **Observers (Push-based, No Recalc)**

Examples:

* `DL.STATUS`
* `DL.LOSS_HISTORY`
* `DL.PREDICT`

**Rules**

* Must never block
* Must not force workbook recalculation
* Must re-emit cached values when nothing changed

**Pattern**

* `ExcelAsyncUtil.Observe`
* Central publish hub (`DlProgressHub`)
* Per-observable caching (Predict keeps last good result; refreshes on publish; X edits change the range key)
* Push only on meaningful state transitions

* * *

3. Fast Start Checklist (Do Exactly This)
-----------------------------------------

1. **Close all Excel.exe processes**
2. Restore: `dotnet restore ExcelDlPlayground/ExcelDlPlayground.csproj -r win-x64`
3. Build: `dotnet build ExcelDlPlayground/ExcelDlPlayground.csproj -c Debug`
4. Load in **64-bit Excel only**: `bin\Debug\net48\ExcelDlPlayground-AddIn64.xll` (or press F5 — Debug profile already launches EXCEL.EXE)
5. Verify natives exist beside the XLL: `torch_cpu.dll`, `LibTorchSharp.dll`
6. In Excel: `=DL.TORCH_TEST_DETAIL()` → must return “torch ok: …”
7. If build fails → **Excel is still running and locking the XLL**

* * *

4. Environment (Fixed & Known-Good)
------------------------------------

* Excel: Microsoft 365 **64-bit**
* Target Framework: **net48**
* PlatformTarget: **x64**
* Prefer32Bit: **false**

### NuGet

* `ExcelDna.AddIn` **1.9.0**
* `TorchSharp-cpu` **0.105.2**
* `libtorch-cpu-win-x64` **2.7.1.0**

### Required References

* `Microsoft.CSharp` (for dynamic TorchSharp calls)
* `System.IO.Compression`

* * *

5. TorchSharp Integration (Hard-Won Lessons)
--------------------------------------------

### Native loading (non-negotiable)

* `torch_cpu.dll` and `LibTorchSharp.dll` **must sit next to the XLL**
* Copied by MSBuild target `CopyTorchNativeBinaries`
* **Never rely on global PATH alone**

### Preload sequence (critical)

```
EnsureTorch():
  - Add baseDir to PATH
  - Set TORCHSHARP_HOME
  - LoadLibrary(LibTorchSharp.dll)
  - LoadLibrary(torch_cpu.dll)
```

Incorrect order causes:

* `TypeInitializationException`
* Silent hangs
* Excel crashes

* * *

6. Session Identity Pattern (Solved)
------------------------------------

### Problem

Excel recalculates everything on workbook open → models recreated

### Solution

**Process-stable session ID**

```
private static readonly string _sessionId = Guid.NewGuid().ToString("N");

[ExcelFunction(Name = "DL.SESSION_ID", IsVolatile = false)]
public static string SessionId() => _sessionId;
```

### Usage

`=DL.MODEL_CREATE("xor:in=2,hidden=8,out=1", DL.SESSION_ID())`  
(Second parameter is optional trigger; same caller + same trigger reuses the cached model id.)

### Behavior

| Event | Result |
| --- | --- |
| Recalc | same model |
| Workbook open | same model |
| Excel restart | new model |

This pattern is **intentional and correct**.

* * *

7. Canonical Worksheet Pattern (Reference)
------------------------------------------

```
A2:B4   → X (features)
C2:C4   → y (labels)

E2: =DL.MODEL_CREATE("xor:in=2,hidden=8,out=1", DL.SESSION_ID())

Z1: 1   (manual trigger cell)
AA1: =DL.TRIGGER_KEY(Z1)

E4: =DL.TRAIN(E2, A2:B4, C2:C4, "epochs=200", Z1)

E8: =DL.LOSS_HISTORY(E2)
E20:=DL.STATUS(E2)

G2: =DL.PREDICT(E2, A2:B4)   (push-based, cached until training completes or X changes)
```

* * *

8. Trigger Semantics (Exact)
-----------------------------

### TriggerKey rules

* Scalars normalized
* Single-cell ranges coerced
* References dereferenced
* Stringified comparison

### Behavior

* **Same trigger → no-op**
* **Different trigger → new run**
* Returned value explains skip reason

* * *

9. Progress & Observables (Why This Works)
-------------------------------------------

### Central Hub

`DlProgressHub.Publish(modelId)`

### Used by:

* `DL.STATUS`
* `DL.LOSS_HISTORY`
* `DL.PREDICT` (refreshes cached prediction on publish; uses range key to rebuild when X edits)

### Publishing cadence

* Training start
* Epoch 1
* Every 5 epochs
* Final epoch
* Completion

### Why no recalculation storm

* Observers are push-based
* Only volatile inspectors use throttled `xlcCalculateNow`
* Workbook-wide recalc is **never forced**

* * *

10. Locks & Threading (Do Not Deviate)
--------------------------------------

### Golden rules

* **Never block Excel UI**
* **Never wait synchronously in observers**
* **Training holds the lock**
* **Prediction never waits**

### Correct patterns

* `TrainLock.WaitAsync()` inside `RunTask`
* `TrainLock.Wait(0)` for prediction (returns cached result if busy)
* Cached fallback if lock unavailable

* * *

11. Debug Helpers (Use These)
------------------------------

* Torch: `DL.TORCH_TEST`, `DL.TORCH_TEST_DETAIL`, `DL.TORCH_NATIVE_CHECK`
* Logging: `DL.LOG_PATH`, `DL.LOG_WRITE_TEST` (log at `bin\Debug\net48\ExcelDlPlayground.log`, fallback `%TEMP%`)

* * *

12. Known Good / Known Bad
---------------------------

### ✅ Proven to work

* Observable-based UI updates
* Session-scoped model identity (caller + trigger cache)
* Cached predictions with push refresh
* Trigger-guarded training
* Throttled recalculation

### ❌ Do NOT do

* Workbook-wide recalc loops
* Blocking `Wait()` on UI thread
* Volatile session IDs
* Recreating models in pure UDFs
* 32-bit Excel
* GPU expectations (CPU-only)

* * *

13. Current Capabilities (Stable)
---------------------------------

* MLP creation & training
* Push-based status & loss curves
* Cached prediction updates
* Weight / gradient / activation inspection
* Save / load (`.dlzip`)
* Robust trigger semantics
* Excel-safe async execution

* * *

14. Next Safe Extensions
------------------------

* `DL.MODEL_INFO`
* Model reset / dispose
* Additional optimizers
* More loss functions
* Structured examples workbook

* * *

Final Note
----------

This architecture works **because it respects Excel’s rules instead of fighting them**.

If something feels “too careful,” it probably is — and it’s there because Excel crashes otherwise.

**Do not simplify without understanding the pattern it replaces.**