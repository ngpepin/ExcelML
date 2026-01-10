# UML Overview

This document captures the core architecture of `ExcelDlPlayground` using PlantUML. The project is a .NET Framework 4.8 Excel-DNA add-in that hosts TorchSharp-based deep learning helpers exposed as Excel UDFs.

Key architectural facts reflected below:

* **`DL.TRAIN` runs async** via `ExcelAsyncUtil.RunTask(...)` and updates in-memory model state.

* **Push-based UI updates** are done via **`ExcelAsyncUtil.Observe(...)` + `DlProgressHub.Publish(...)`** (STATUS / LOSS\_HISTORY / PREDICT).

* `QueueRecalcOnce(xlcCalculateNow)` is a **throttled fallback**, mainly useful for volatile/legacy cells — _not required_ for the observable functions.

* `DL.PREDICT` is implemented as an **observable with caching** so it doesn’t flicker or block Excel; it recomputes only when a **new stable model version** is available.

- - -

<div style="page-break-after: always;"></div>

## Component Diagram

### Notes

- Removed “xlcCalculateNow on load” as a *hard* behavior — startup recalc is not fundamental to the architecture.

- Made the **observable loop explicit**: `DlProgressHub` signals, and `ExcelAsyncUtil.Observe(...)` consumers re-emit values.

- Clarified that **Torch** includes both managed + native dependencies.

```plantuml
@startuml
title ExcelDlPlayground - Component Overview
skinparam componentStyle rectangle
skinparam packageStyle rectangle

package "Excel Host" {
  [Excel UI] as ExcelUI
  [Excel Calc Engine] as ExcelCalc
  [Excel Macro Queue] as ExcelMacro
}

package "Excel-DNA Add-in (net48)" {
  [AddInStartup] as Startup
  [DlRibbon] as Ribbon
  [Functions] as BasicUdfs
  [AsyncFunctions] as AsyncUdfs
  [DlFunctions] as DlUdfs
  [DlProgressHub] as Hub
}

package "DL Core (in-memory)" {
  [DlRegistry] as Registry
  [DlModelState] as ModelState
}

node "TorchSharp / LibTorch (native + managed)" as Torch
node "File System" as FS

' --- UI + call surfaces ---
ExcelUI --> Ribbon : Ribbon callbacks
ExcelCalc --> DlUdfs : UDF calls (DL.*)
ExcelCalc --> BasicUdfs : UDF calls (non-DL)
ExcelCalc --> AsyncUdfs : UDF calls (async demos)

' --- Startup ---
Startup --> ExcelMacro : (optional) queue recalc / init hooks
Startup ..> Torch : ensure native deps discoverable (via EnsureTorch)

' --- Basic + async helpers (non-DL) ---
BasicUdfs --> ExcelCalc : simple UDFs
AsyncUdfs ..> ExcelCalc : async demos / wrappers

' --- DL surface ---
DlUdfs --> Registry : create/load/save/get model state
DlUdfs --> Torch : tensors, autograd, modules, optimizers
DlUdfs --> FS : save/load model packages (.dlzip)

' --- Push updates to Excel ---
DlUdfs --> Hub : Publish(modelId) during training + completion
Hub --> ExcelCalc : Notifies registered IExcelObserver instances\n(used by Observe(...) UDFs)

' --- Registry owns model instances ---
Registry o-- ModelState : manages instances keyed by modelId

@enduml
```

<div style="page-break-after: always;"></div>

## Class Diagram (Key Types)

### Notes

- `Status`, `LossHistory`, **and `Predict`** are modelled as *observable outputs* (this is crucial).

- `QueueRecalcOnce` is explicitly framed as a **fallback**, not the primary update mechanism.

- `PredictObservable` caching behaviour is captured at a conceptual level.

<div style="page-break-after: always;"></div>

```plantuml
@startuml
skinparam classAttributeIconSize 0
skinparam shadowing false

class StatusObservable
class LossObservable
class PredictObservable {
  - _lastVersion : long
  - _lastResult : object
  - ComputeOrGetCached()
}
class DlFunctions {
  .. Torch init / native preload ..
  - bool _torchInitialized
  - string[] TorchNativeFiles
  - string LogPathSafe
  - EnsureTorch()
  - GetTorchBaseDir()
  - GetMissingTorchNativeFiles()

  .. Recalc helper (fallback only) ..
  - QueueRecalcOnce(reason, force)

  .. Public UDFs ..
  + ModelCreate(description, trigger = null)
  + Train(modelId, X, y, opts, trigger)
  + Status(modelId) : IObservable (via Observe)
  + LossHistory(modelId) : IObservable (via Observe)
  + Predict(modelId, X) : IObservable (via Observe + caching)

  + Weights(modelId, layer) : volatile inspector
  + Grads(modelId, layer) : volatile inspector
  + Activations(modelId, X, layer) : volatile inspector

  + Save(modelId, path)
  + Load(path)

  .. Builders for observables ..
  - BuildStatus(modelId)
  - BuildLossTable(modelId)

  .. Predict helpers ..
  - (PredictObservable caches lastResult)
  - (Predict recompute only when TrainingVersion changes AND IsTraining=false)

  .. Helpers ..
  - ParseIntOpt()
  - ParseDoubleOpt()
  - ParseStringOpt()
  - ParseBoolOpt()
  - BuildTensorFromRange()
  - ResolveLayerName()
  - RunForwardActivations()
  - TensorToObjectArray()
  - BuildWeightMatrix()
  - TriggerKey()
  - BuildDefaultMlp()
  - CreateOptimizer()
  - SaveStateDict()
  - Log(message)
}

class DlProgressHub {
  - _subs : ConcurrentDictionary<string, HashSet<IExcelObserver>>
  + Subscribe(modelId, observer) : IDisposable
  + Publish(modelId) : void
}

class DlRegistry {
  - _models : ConcurrentDictionary<string, DlModelState>
  + CreateModel(description) : string
  + TryGet(modelId, out state) : bool
  + Upsert(modelId, state) : void
}

class DlModelState {
  + Description : string
  + LastTriggerKey : string

  + TrainLock : SemaphoreSlim
  + IsTraining : bool
  + TrainingVersion : long
  + LastEpoch : int
  + LastLoss : double

  + LossHistory : List<(epoch,loss)>

  + TorchModel : Module
  + Optimizer : Optimizer
  + LossFn : Module
  + OptimizerName : string
  + LearningRate : double

  + InputDim : int
  + HiddenDim : int
  + OutputDim : int

  + WeightSnapshot : Dictionary<string, LayerTensorSnapshot>
  + GradSnapshot : Dictionary<string, LayerTensorSnapshot>
  + ActivationSnapshot : Dictionary<string, Tensor>

  + UpdateWeightSnapshot()
  + UpdateGradSnapshot()
  + UpdateActivationSnapshot()
}

class LayerTensorSnapshot {
  + Weight : Tensor
  + Bias : Tensor
  + Dispose()
}

class AddInStartup {
  + AutoOpen()
  + AutoClose()
}
class DlRibbon { 
  + GetCustomUI(id) 
  + OnLoad(ui) 
  + OnHelloClick(control) 
  + OnInvalidateClick(control) 
}
class Functions {
 + SayHello(name) 
 + MatMul(a,b) 
}
class AsyncFunctions { 
  + WaitAsync(ms) 
}

DlFunctions --> DlRegistry
DlFunctions --> DlProgressHub : Publish + Observe subscribers
DlRegistry o-- DlModelState
DlModelState o-- LayerTensorSnapshot

DlFunctions --> "TorchSharp" : uses
DlFunctions --> "ExcelDna.Integration" : UDF surface
DlFunctions --> "ExcelAsyncUtil" : RunTask / Observe
AddInStartup --> "Excel host" : add-in lifecycle

DlFunctions ..> StatusObservable : Observe(...)
DlFunctions ..> LossObservable : Observe(...)
DlFunctions ..> PredictObservable : Observe(...)\n(cached)

StatusObservable ..> DlProgressHub : Subscribe(modelId)
LossObservable ..> DlProgressHub : Subscribe(modelId)
PredictObservable ..> DlProgressHub : Subscribe(modelId)

DlProgressHub ..> "IExcelObserver" : OnNext(modelId)

@enduml
```

<div style="page-break-after: always;"></div>

## Sequence (Training happy path)

### Notes

- Shows that **PREDICT recompute happens at “stable model” time** (when `IsTraining=false` and `TrainingVersion` changed).

- Makes the observer loop explicit, rather than implying calc engine pulls values.

- Explicitly marks macro recalc as optional.

```plantuml
@startuml
title DL.TRAIN + push updates (STATUS / LOSS / PREDICT)

actor User
participant Excel
participant "DlFunctions.Train\n(ExcelAsyncUtil.RunTask)" as Train
participant DlRegistry as Registry
participant DlModelState as Model
participant DlProgressHub as Hub
participant "Excel-DNA Observe(...) pipeline\n(IExcelObservable/IExcelObserver)" as Obs
participant TorchSharp as Torch
participant "Excel Macro Queue\n(optional throttled recalc)" as MacroQ

User -> Excel : Change trigger cell
Excel -> Train : DL.TRAIN(modelId, X, y, opts, trigger)

Train -> Registry : TryGet(modelId)
Registry --> Train : ModelState

Train -> Model : TrainLock.WaitAsync()
Train -> Torch : EnsureTorch()
Train -> Model : IsTraining=true\nTrainingVersion++

Train -> Hub : Publish(modelId)  // "training started"
Hub -> Obs : OnNext(modelId)\nObservers call builders; Excel updates cell values

Train -> Torch : BuildTensorFromRange(X,y)
loop epochs (1..N)
  Train -> Torch : forward(x)
  Train -> Torch : loss(output,y)
  Train -> Torch : backward()
  Train -> Torch : optimizer.step()
  Train -> Model : LastEpoch/LastLoss updated\nLossHistory.Add(...)
  Train -> Hub : Publish(modelId)  // periodic progress (e.g. epoch 1, every 5, final)
  Hub -> Obs : OnNext(modelId)\nSTATUS/LOSS update\nPREDICT may stay cached while training
end

Train -> Model : UpdateWeightSnapshot()
Train -> Model : LastTriggerKey = TriggerKey(trigger)
Train -> Model : IsTraining=false

Train -> Hub : Publish(modelId)  // "training completed"
Hub -> Obs : OnNext(modelId)\nPREDICT recomputes (IsTraining=false && version changed)

Train -> MacroQ : QueueRecalcOnce(train-complete)\n(optional fallback)
Train -> Model : TrainLock.Release()
Train --> Excel : return {status=done, epochs, final_loss}

@enduml
```