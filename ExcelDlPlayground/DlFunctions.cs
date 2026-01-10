using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using TorchSharp;
using static TorchSharp.torch;
using System.Runtime.InteropServices;
using System.IO.Compression;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using TorchSharp.Modules;

/// <summary>
/// Excel-DNA UDF surface and supporting helpers for TorchSharp-based deep learning inside Excel.
/// </summary>
public static class DlFunctions
{
    private static bool _torchInitialized;
    private static readonly string[] TorchNativeFiles = { "LibTorchSharp.dll", "torch_cpu.dll" };

    private static string LogPathSafe
    {
        get
        {
            try
            {
                return Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? AppDomain.CurrentDomain.BaseDirectory, "ExcelDlPlayground.log");
            }
            catch
            {
                return Path.Combine(Path.GetTempPath(), "ExcelDlPlayground.log");
            }
        }
    }

    /// <summary>
    /// Ensures TorchSharp native binaries are discoverable and preloaded once per process.
    /// </summary>
    private static void EnsureTorch()
    {
        if (_torchInitialized) return;

        var baseDir = GetTorchBaseDir();

        var path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        if (!path.Split(';').Any(p => string.Equals(p, baseDir, StringComparison.OrdinalIgnoreCase)))
        {
            Environment.SetEnvironmentVariable("PATH", baseDir + ";" + path);
        }

        Environment.SetEnvironmentVariable("TORCHSHARP_HOME", baseDir);

        // Load native shim first
        var libTorchSharpPath = Path.Combine(baseDir, "LibTorchSharp.dll");
        if (File.Exists(libTorchSharpPath))
        {
            var handleShim = LoadLibrary(libTorchSharpPath);
            if (handleShim != IntPtr.Zero)
            {
                Log($"EnsureTorch loaded {libTorchSharpPath}");
            }
            else
            {
                Log($"EnsureTorch failed LoadLibrary {libTorchSharpPath}, GetLastError={Marshal.GetLastWin32Error()}");
            }
        }
        else
        {
            Log($"EnsureTorch: LibTorchSharp.dll not found in {baseDir}");
        }

        var torchCpuPath = Path.Combine(baseDir, "torch_cpu.dll");
        if (File.Exists(torchCpuPath))
        {
            try
            {
                var handle = LoadLibrary(torchCpuPath);
                if (handle != IntPtr.Zero)
                {
                    Log($"EnsureTorch loaded {torchCpuPath}");
                    _torchInitialized = true;
                }
                else
                {
                    Log($"EnsureTorch failed LoadLibrary {torchCpuPath}, GetLastError={Marshal.GetLastWin32Error()}");
                }
            }
            catch (Exception ex)
            {
                Log($"EnsureTorch exception loading {torchCpuPath}: {ex}");
            }
        }
        else
        {
            Log($"EnsureTorch: torch_cpu.dll not found in {baseDir}");
        }
    }

    /// <summary>
    /// Resolves the base directory where TorchSharp assemblies and natives are expected.
    /// </summary>
    private static string GetTorchBaseDir()
    {
        string baseDir = null;
        try
        {
            var loc = Assembly.GetExecutingAssembly().Location;
            if (!string.IsNullOrWhiteSpace(loc))
                baseDir = Path.GetDirectoryName(loc);
        }
        catch { }
        if (string.IsNullOrWhiteSpace(baseDir))
            baseDir = AppDomain.CurrentDomain.BaseDirectory;

        return baseDir;
    }

    /// <summary>
    /// Returns a list of TorchSharp native files missing from the expected base directory.
    /// </summary>
    private static List<string> GetMissingTorchNativeFiles(string baseDir)
    {
        var missing = new List<string>();
        foreach (var file in TorchNativeFiles)
        {
            if (!File.Exists(Path.Combine(baseDir, file)))
            {
                missing.Add(file);
            }
        }

        return missing;
    }

    /// <summary>
    /// Creates a model entry and initializes a small MLP with defaults based on description text.
    /// </summary>
    [ExcelFunction(Name = "DL.MODEL_CREATE", Description = "Create a model and return a model_id" /* non-volatile */)]
    public static object ModelCreate(string description)
    {
        try
        {
            EnsureTorch();
            // 1) Create registry entry
            var id = DlRegistry.CreateModel(description ?? "");
            if (!DlRegistry.TryGet(id, out var model))
                return "#ERR: registry failure";

            // 2) Parse a few knobs from description (very lightweight)
            // Supported examples:
            //  "xor:in=2,hidden=8,out=1"
            //  "in=4,hidden=16,out=1"
            int input = ParseIntOpt(description, "in", 2);
            int hidden = ParseIntOpt(description, "hidden", 8);
            int output = ParseIntOpt(description, "out", 1);

            model.InputDim = input;
            model.HiddenDim = hidden;
            model.OutputDim = output;

            // 3) Build a tiny MLP: Linear -> Tanh -> Linear
            // NOTE: For XOR/binary classification we'll use BCEWithLogitsLoss (so last layer is raw logits).
            var net = BuildDefaultMlp(input, hidden, output);

            model.TorchModel = net;
            model.LossFn = torch.nn.BCEWithLogitsLoss(); // good default for XOR
            model.OptimizerName = "adam";
            model.LearningRate = 0.1;
            model.Optimizer = CreateOptimizer(model);

            Log($"ModelCreate | desc={description ?? "<null>"} | id={id} | in={input} hidden={hidden} out={output}");
            return id;
        }
        catch (Exception ex)
        {
            Log($"ModelCreate error | desc={description ?? "<null>"} | ex={ex}");
            return "#ERR: " + ex.Message;
        }
    }

    /// <summary>
    /// Serializes a model and metadata to a zip payload on disk.
    /// </summary>
    [ExcelFunction(Name = "DL.SAVE", Description = "Save a model to disk")]
    public static object Save(string model_id, string path)
    {
        if (string.IsNullOrWhiteSpace(model_id))
            return "#ERR: model_id is required";

        if (string.IsNullOrWhiteSpace(path))
            return "#ERR: path is required";

        if (!DlRegistry.TryGet(model_id, out var model))
            return "#MODEL! Unknown model_id";

        if (model.TorchModel == null)
            return "#ERR: model not initialized. Call DL.MODEL_CREATE first.";

        EnsureTorch();

        var fullPath = Path.GetFullPath(path);
        var tempDir = Path.Combine(Path.GetTempPath(), "ExcelDlPlayground", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        model.TrainLock.Wait();
        try
        {
            var meta = new DlModelPersistence
            {
                FormatVersion = 1,
                ModelId = model_id,
                Description = model.Description ?? string.Empty,
                InputDim = model.InputDim,
                HiddenDim = model.HiddenDim,
                OutputDim = model.OutputDim,
                OptimizerName = model.OptimizerName ?? "adam",
                LearningRate = model.LearningRate,
                LastTriggerKey = model.LastTriggerKey,
                LossHistory = model.LossHistory.Select(entry => new DlLossEntry { Epoch = entry.epoch, Loss = entry.loss }).ToList()
            };

            var weightsPath = Path.Combine(tempDir, "model.pt");
            var metadataPath = Path.Combine(tempDir, "metadata.json");

            var stateDict = model.TorchModel.state_dict();
            SaveStateDict(stateDict, weightsPath);
            WriteMetadata(metadataPath, meta);

            var outputDir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(outputDir))
                Directory.CreateDirectory(outputDir);

            using (var fileStream = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
            using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Create))
            {
                AddFileToArchive(archive, weightsPath, "model.pt");
                AddFileToArchive(archive, metadataPath, "metadata.json");
            }

            Log($"ModelSave | id={model_id} | path={fullPath}");
            return "saved: " + fullPath;
        }
        catch (Exception ex)
        {
            Log($"ModelSave error | id={model_id} | path={fullPath} | ex={ex}");
            return "#ERR: " + ex.Message;
        }
        finally
        {
            model.TrainLock.Release();
            try
            {
                if (Directory.Exists(tempDir))
                    Directory.Delete(tempDir, recursive: true);
            }
            catch
            {
            }
        }
    }

    /// <summary>
    /// Loads a serialized model package (.dlzip) from disk and rehydrates registry state.
    /// </summary>
    [ExcelFunction(Name = "DL.LOAD", Description = "Load a model from disk")]
    public static object Load(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return "#ERR: path is required";

        var fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
            return "#ERR: file not found";

        EnsureTorch();

        var tempDir = Path.Combine(Path.GetTempPath(), "ExcelDlPlayground", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            DlModelPersistence meta;
            var weightsPath = Path.Combine(tempDir, "model.pt");

            using (var fileStream = new FileStream(fullPath, FileMode.Open, FileAccess.Read))
            using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Read))
            {
                var metadataEntry = archive.GetEntry("metadata.json");
                var modelEntry = archive.GetEntry("model.pt");
                if (metadataEntry == null || modelEntry == null)
                    return "#ERR: invalid model package (missing entries)";

                using (var metadataStream = metadataEntry.Open())
                {
                    meta = ReadMetadata(metadataStream);
                }

                using (var modelStream = modelEntry.Open())
                using (var output = new FileStream(weightsPath, FileMode.Create, FileAccess.Write))
                {
                    modelStream.CopyTo(output);
                }
            }

            if (meta == null || string.IsNullOrWhiteSpace(meta.ModelId))
                return "#ERR: invalid metadata";

            if (meta.FormatVersion != 1)
                return "#ERR: unsupported model format";

            if (meta.InputDim <= 0 || meta.HiddenDim <= 0 || meta.OutputDim <= 0)
                return "#ERR: invalid model dimensions";

            var model = new DlModelState(meta.Description ?? string.Empty)
            {
                InputDim = meta.InputDim,
                HiddenDim = meta.HiddenDim,
                OutputDim = meta.OutputDim,
                OptimizerName = string.IsNullOrWhiteSpace(meta.OptimizerName) ? "adam" : meta.OptimizerName,
                LearningRate = meta.LearningRate,
                LastTriggerKey = meta.LastTriggerKey
            };

            model.TorchModel = BuildDefaultMlp(model.InputDim, model.HiddenDim, model.OutputDim);
            model.LossFn = torch.nn.BCEWithLogitsLoss();
            model.Optimizer = CreateOptimizer(model);

            var stateDict = torch.load(weightsPath) as IDictionary<string, Tensor>;
            if (stateDict == null)
                return "#ERR: invalid weights payload";

            var dict = stateDict as Dictionary<string, Tensor> ?? new Dictionary<string, Tensor>(stateDict);

            model.TorchModel.load_state_dict(dict);
            model.LossHistory.Clear();
            if (meta.LossHistory != null)
            {
                foreach (var entry in meta.LossHistory)
                {
                    model.LossHistory.Add((entry.Epoch, entry.Loss));
                }
            }

            model.UpdateWeightSnapshot();
            DlRegistry.Upsert(meta.ModelId, model);

            Log($"ModelLoad | id={meta.ModelId} | path={fullPath}");
            return meta.ModelId;
        }
        catch (Exception ex)
        {
            Log($"ModelLoad error | path={fullPath} | ex={ex}");
            return "#ERR: " + ex.Message;
        }
        finally
        {
            try
            {
                if (Directory.Exists(tempDir))
                    Directory.Delete(tempDir, recursive: true);
            }
            catch
            {
            }
        }
    }

    // Throttled recalc helper to avoid storms (no workbook-wide force)
    private static volatile bool _recalcQueued;
    private static volatile bool _recalcPending;

    /// <summary>
    /// Queues a single recalculation macro, coalescing duplicate requests to avoid storms.
    /// </summary>
    private static void QueueRecalcOnce(string reason, bool force)
    {
        if (_recalcQueued)
        {
            _recalcPending = true; // coalesce concurrent requests (progress or forced)
            return;
        }

        _recalcQueued = true;
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    XlCall.Excel(XlCall.xlcCalculateNow);
                }
                catch { }
                finally
                {
                    _recalcQueued = false;
                    if (_recalcPending)
                    {
                        _recalcPending = false;
                        QueueRecalcOnce("pending", false);
                    }
                }
            });
        }
        catch
        {
            _recalcQueued = false;
        }
    }

    /// <summary>
    /// Returns a snapshot of training state for a given model.
    /// </summary>
    [ExcelFunction(Name = "DL.STATUS", Description = "Show training status for a model", IsVolatile = true)]
    public static object Status(string model_id)
    {
        if (!DlRegistry.TryGet(model_id, out var model))
            return "#MODEL! Unknown model_id";

        return new object[,]
        {
            { "model", model_id },
            { "status", model.IsTraining ? "training" : "idle" },
            { "last_epoch", model.LastEpoch },
            { "last_loss", double.IsNaN(model.LastLoss) ? "" : model.LastLoss.ToString("G6", CultureInfo.InvariantCulture) },
            { "last_trigger", model.LastTriggerKey ?? "<null>" },
            { "version", model.TrainingVersion }
        };
    }

    /// <summary>
    /// Trains the specified model asynchronously using provided feature and label ranges.
    /// </summary>
    [ExcelFunction(Name = "DL.TRAIN", Description = "Train a model (triggered) and return summary")]
    public static object Train(string model_id, object[,] X, object[,] y, string opts, object trigger)
    {
        var key = TriggerKey(trigger);

        if (!DlRegistry.TryGet(model_id, out var model))
            return "#MODEL! Unknown model_id";

        Log($"Train enter | model={model_id} | key={key} | last={model.LastTriggerKey ?? "<null>"}");

        if (model.LastTriggerKey == key)
        {
            return new object[,] { { "skipped", "trigger unchanged (set trigger cell to a new value to retrain)" }, { "last", model.LastTriggerKey ?? "<null>" }, { "curr", key } };
        }

        // Fast-path guard: if another training run is in-progress, return immediately instead of waiting (avoids long #N/A).
        if (!model.TrainLock.Wait(0))
        {
            return new object[,]
            {
                { "busy", "training in progress" },
                { "hint", "retry after current training completes" }
            };
        }
        model.TrainLock.Release();

        var functionName = nameof(Train);
        var parameters = new object[] { model_id, opts ?? "", key };

        return ExcelAsyncUtil.RunTask<object>(functionName, parameters, async () =>
         {
            await model.TrainLock.WaitAsync().ConfigureAwait(false);
            try
            {
                Log($"Train lock acquired | model={model_id} | key={key} | last={model.LastTriggerKey ?? "<null>"}");

                if (model.LastTriggerKey == key)
                {
                    Log($"Train early no-op inside lock | model={model_id} | key={key} | last={model.LastTriggerKey ?? "<null>"}");
                    return new object[,] { { "skipped", "trigger unchanged (set trigger cell to a new value to retrain)" }, { "last", model.LastTriggerKey ?? "<null>" }, { "curr", key } };
                }

                int epochs = ParseIntOpt(opts, "epochs", 20);
                double learningRate = ParseDoubleOpt(opts, "lr", 0.1);
                bool resetModel = ParseBoolOpt(opts, "reset", false);
                EnsureTorch();

                if (model.TorchModel == null)
                {
                    return "#ERR: model not initialized. Call DL.MODEL_CREATE first.";
                }

                model.IsTraining = true;
                model.TrainingVersion++;
                model.LastEpoch = 0;
                model.LastLoss = double.NaN;

                if (resetModel)
                {
                    if (model.Optimizer is IDisposable oldOpt)
                    {
                        oldOpt.Dispose();
                    }

                    model.TorchModel = BuildDefaultMlp(model.InputDim, model.HiddenDim, model.OutputDim);
                    model.LossFn = torch.nn.BCEWithLogitsLoss();
                    model.Optimizer = null;
                    model.WeightSnapshot.Clear();
                    model.GradSnapshot.Clear();
                    model.LossHistory.Clear();
                }

                if (model.LossFn == null)
                {
                    model.LossFn = torch.nn.BCEWithLogitsLoss();
                }

                var optimizerName = ParseStringOpt(opts, "optim", model.OptimizerName ?? "adam");
                if (string.IsNullOrWhiteSpace(optimizerName))
                    optimizerName = "adam";
                optimizerName = optimizerName.Trim().ToLowerInvariant();

                if (optimizerName != "adam" && optimizerName != "sgd")
                    return $"#ERR: unsupported optimizer '{optimizerName}'. Use optim=adam or optim=sgd.";

                bool optimizerNeedsReset = model.Optimizer == null
                    || !string.Equals(model.OptimizerName, optimizerName, StringComparison.OrdinalIgnoreCase)
                    || Math.Abs(model.LearningRate - learningRate) > 1e-12;

                if (resetModel || optimizerNeedsReset)
                {
                    if (model.Optimizer is IDisposable disposable)
                    {
                        disposable.Dispose();
                    }

                    model.OptimizerName = optimizerName;
                    model.LearningRate = learningRate;
                    model.Optimizer = CreateOptimizer(model);
                }

                model.LossHistory.Clear();
                double loss = 0.0;

                using (var xTensor = BuildTensorFromRange(X, model.InputDim, "X"))
                using (var yTensor = BuildTensorFromRange(y, model.OutputDim, "y"))
                {
                    model.TorchModel.train();

                    for (int e = 1; e <= epochs; e++)
                    {
                        model.Optimizer.zero_grad();
                        using (var output = (Tensor)((dynamic)model.TorchModel).forward(xTensor))
                        using (var lossTensor = (Tensor)((dynamic)model.LossFn).forward(output, yTensor))
                        {
                            lossTensor.backward();
                            model.UpdateGradSnapshot();
                            model.Optimizer.step();
                            loss = lossTensor.ToSingle();
                        }

                        model.LossHistory.Add((e, loss));
                        model.LastEpoch = e;
                        model.LastLoss = loss;

                        // queue a recalc each epoch (throttled in QueueRecalcOnce)
                        QueueRecalcOnce("loss-progress", false);
                        await Task.Delay(1).ConfigureAwait(false);
                    }

                    model.UpdateWeightSnapshot();
                }

                model.LastTriggerKey = key;
                model.IsTraining = false;
                QueueRecalcOnce("train-complete", true);
                Log($"Train complete | model={model_id} | key set to {key} | epochs={epochs} | final_loss={loss}");

                return new object[,]
                {
                { "status", "done" },
                { "epochs", epochs },
                { "final_loss", loss.ToString("G6", CultureInfo.InvariantCulture) }
                };
            }
            finally
            {
                model.IsTraining = false;
                model.TrainLock.Release();
            }
        });
    }

    /// <summary>
    /// Returns epoch/loss pairs recorded during the last training run.
    /// </summary>
    [ExcelFunction(Name = "DL.LOSS_HISTORY", Description = "Spill epoch/loss history for a model", IsVolatile = true)]
    public static object LossHistory(string model_id)
    {
        if (!DlRegistry.TryGet(model_id, out var model))
            return "#MODEL! Unknown model_id";

        if (model.LossHistory.Count == 0)
            return new object[,] { { "empty", "no training history" } };

        var n = model.LossHistory.Count;
        var output = new object[n + 1, 2];
        output[0, 0] = "epoch";
        output[0, 1] = "loss";

        for (int i = 0; i < n; i++)
        {
            output[i + 1, 0] = model.LossHistory[i].epoch;
            output[i + 1, 1] = model.LossHistory[i].loss;
        }
        return output;
    }

    /// <summary>
    /// Runs inference for the given feature range and returns outputs as a spilled array.
    /// </summary>
    [ExcelFunction(Name = "DL.PREDICT", Description = "Predict outputs for X (spilled)")]
    public static object Predict(string model_id, object[,] X)
    {
        if (!DlRegistry.TryGet(model_id, out var model))
            return "#MODEL! Unknown model_id";

        if (model.TorchModel == null)
            return "#ERR: model not initialized. Call DL.MODEL_CREATE first.";

        EnsureTorch();

        model.TrainLock.Wait();
        try
        {
            using (var xTensor = BuildTensorFromRange(X, model.InputDim, "X"))
            using (torch.no_grad())
            {
                model.TorchModel.eval();
                using (var output = (Tensor)((dynamic)model.TorchModel).forward(xTensor))
                using (var outputCpu = output.detach().cpu())
                {
                    return TensorToObjectArray(outputCpu);
                }
            }
        }
        finally
        {
            model.TrainLock.Release();
        }
    }

    /// <summary>
    /// Returns a snapshot of weights for the specified layer as a spilled matrix.
    /// </summary>
    [ExcelFunction(Name = "DL.WEIGHTS", Description = "Inspect weights for a layer (spilled)", IsVolatile = true)]
    public static object Weights(string model_id, object layer)
    {
        if (!DlRegistry.TryGet(model_id, out var model))
            return "#MODEL! Unknown model_id";

        var layerName = ResolveLayerName(model, layer, requireWeightedLayer: true, out var error);
        if (error != null)
            return error;

        if (!model.WeightSnapshot.TryGetValue(layerName, out var snapshot))
            return "#ERR: no weight snapshot. Train the model first.";

        return BuildWeightMatrix(snapshot.Weight, snapshot.Bias);
    }

    /// <summary>
    /// Returns a snapshot of gradients for the specified layer as a spilled matrix.
    /// </summary>
    [ExcelFunction(Name = "DL.GRADS", Description = "Inspect gradients for a layer (spilled)", IsVolatile = true)]
    public static object Grads(string model_id, object layer)
    {
        if (!DlRegistry.TryGet(model_id, out var model))
            return "#MODEL! Unknown model_id";

        var layerName = ResolveLayerName(model, layer, requireWeightedLayer: true, out var error);
        if (error != null)
            return error;

        if (!model.GradSnapshot.TryGetValue(layerName, out var snapshot))
            return "#ERR: no gradient snapshot. Train the model first.";

        return BuildWeightMatrix(snapshot.Weight, snapshot.Bias);
    }

    /// <summary>
    /// Runs a forward pass to capture activations for a given layer and feature set.
    /// </summary>
    [ExcelFunction(Name = "DL.ACTIVATIONS", Description = "Inspect activations for a layer given X (spilled)", IsVolatile = true)]
    public static object Activations(string model_id, object[,] X, object layer)
    {
        if (!DlRegistry.TryGet(model_id, out var model))
            return "#MODEL! Unknown model_id";

        if (model.TorchModel == null)
            return "#ERR: model not initialized. Call DL.MODEL_CREATE first.";

        var layerName = ResolveLayerName(model, layer, requireWeightedLayer: false, out var error);
        if (error != null)
            return error;

        EnsureTorch();

        model.TrainLock.Wait();
        try
        {
            using (var xTensor = BuildTensorFromRange(X, model.InputDim, "X"))
            using (torch.no_grad())
            {
                var activations = RunForwardActivations(model, xTensor);
                model.UpdateActivationSnapshot(activations);
                if (!activations.TryGetValue(layerName, out var activation))
                    return "#ERR: layer not found";

                return TensorToObjectArray(activation);
            }
        }
        finally
        {
            model.TrainLock.Release();
        }
    }

    /// <summary>
    /// Parses an integer option from the opts string, returning a default when absent.
    /// </summary>
    private static int ParseIntOpt(string opts, string key, int defaultValue)
    {
        if (string.IsNullOrWhiteSpace(opts)) return defaultValue;
        var parts = opts.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
        {
            var kv = p.Split('=');
            if (kv.Length == 2 && kv[0].Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(kv[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v))
                    return v;
            }
        }
        return defaultValue;
    }

    /// <summary>
    /// Parses a double option from the opts string, returning a default when absent.
    /// </summary>
    private static double ParseDoubleOpt(string opts, string key, double defaultValue)
    {
        if (string.IsNullOrWhiteSpace(opts)) return defaultValue;
        var parts = opts.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
        {
            var kv = p.Split('=');
            if (kv.Length == 2 && kv[0].Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                if (double.TryParse(kv[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                    return v;
            }
        }
        return defaultValue;
    }

    /// <summary>
    /// Parses a string option from the opts string, returning a default when absent.
    /// </summary>
    private static string ParseStringOpt(string opts, string key, string defaultValue)
    {
        if (string.IsNullOrWhiteSpace(opts)) return defaultValue;
        var parts = opts.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
        {
            var kv = p.Split('=');
            if (kv.Length == 2 && kv[0].Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                return kv[1].Trim();
            }
        }
        return defaultValue;
    }

    /// <summary>
    /// Parses a boolean option from the opts string, returning a default when absent.
    /// </summary>
    private static bool ParseBoolOpt(string opts, string key, bool defaultValue)
    {
        if (string.IsNullOrWhiteSpace(opts)) return defaultValue;
        var parts = opts.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
        {
            var kv = p.Split('=');
            if (kv.Length == 2 && kv[0].Trim().Equals(key, StringComparison.OrdinalIgnoreCase))
            {
                var v = kv[1].Trim();
                if (string.Equals(v, "1") || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase) || string.Equals(v, "yes", StringComparison.OrdinalIgnoreCase))
                    return true;
                if (string.Equals(v, "0") || string.Equals(v, "false", StringComparison.OrdinalIgnoreCase) || string.Equals(v, "no", StringComparison.OrdinalIgnoreCase))
                    return false;
            }
        }
        return defaultValue;
    }

    /// <summary>
    /// Builds a Torch tensor from a 2D Excel range, validating expected column count.
    /// </summary>
    private static Tensor BuildTensorFromRange(object[,] values, int expectedCols, string label)
    {
        if (values == null)
            throw new ArgumentException($"Range {label} is null");

        int rows = values.GetLength(0);
        int cols = values.GetLength(1);

        if (expectedCols > 0 && cols != expectedCols)
            throw new ArgumentException($"Range {label} must have {expectedCols} columns, got {cols}");

        var data = new float[rows * cols];
        int idx = 0;
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                var cell = values[r, c];
                if (cell == null || cell is ExcelEmpty || cell is ExcelMissing)
                {
                    data[idx++] = 0f;
                    continue;
                }

                if (cell is double d)
                {
                    data[idx++] = (float)d;
                    continue;
                }

                if (cell is float f)
                {
                    data[idx++] = f;
                    continue;
                }

                if (double.TryParse(cell.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
                {
                    data[idx++] = (float)parsed;
                    continue;
                }

                throw new ArgumentException($"Range {label} contains non-numeric value at ({r + 1},{c + 1})");
            }
        }

        return torch.tensor(data, new long[] { rows, cols }, dtype: ScalarType.Float32);
    }

    /// <summary>
    /// Resolves a layer name from index or name input, validating against model layers.
    /// </summary>
    private static string ResolveLayerName(DlModelState model, object layer, bool requireWeightedLayer, out string error)
    {
        error = null;
        var normalized = NormalizeScalar(layer);
        var layers = GetLayerNames(model, requireWeightedLayer);

        if (layers.Count == 0)
        {
            error = "#ERR: model has no layers";
            return null;
        }

        if (normalized is double d)
        {
            int index = (int)d;
            if (Math.Abs(d - index) > 1e-9)
            {
                error = "#ERR: layer index must be an integer";
                return null;
            }

            if (index < 1 || index > layers.Count)
            {
                error = $"#ERR: layer index out of range (1-{layers.Count})";
                return null;
            }

            return layers[index - 1];
        }

        if (normalized is int i)
        {
            if (i < 1 || i > layers.Count)
            {
                error = $"#ERR: layer index out of range (1-{layers.Count})";
                return null;
            }

            return layers[i - 1];
        }

        if (normalized is string s && !string.IsNullOrWhiteSpace(s))
        {
            if (layers.Any(name => name.Equals(s, StringComparison.OrdinalIgnoreCase)))
                return layers.First(name => name.Equals(s, StringComparison.OrdinalIgnoreCase));

            error = "#ERR: unknown layer name";
            return null;
        }

        error = "#ERR: invalid layer";
        return null;
    }

    /// <summary>
    /// Enumerates child layers optionally filtering to weighted layers only.
    /// </summary>
    private static List<string> GetLayerNames(DlModelState model, bool requireWeightedLayer)
    {
        var names = new List<string>();
        if (model.TorchModel == null)
            return names;

        foreach (var layer in model.TorchModel.named_children())
        {
            if (requireWeightedLayer && !(layer.module is Linear))
                continue;

            names.Add(layer.name);
        }

        return names;
    }

    /// <summary>
    /// Runs a forward pass capturing intermediate activations for each named child layer.
    /// </summary>
    private static Dictionary<string, Tensor> RunForwardActivations(DlModelState model, Tensor xTensor)
    {
        var activations = new Dictionary<string, Tensor>(StringComparer.OrdinalIgnoreCase);
        var intermediates = new System.Collections.Generic.List<Tensor>();
        var current = xTensor;

        foreach (var layer in model.TorchModel.named_children())
        {
            var output = (Tensor)((dynamic)layer.module).forward(current);
            intermediates.Add(output);
            activations[layer.name] = output.detach().clone().cpu();
            current = output;
        }

        foreach (var tensor in intermediates)
        {
            tensor.Dispose();
        }

        return activations;
    }

    /// <summary>
    /// Converts a Torch tensor into a 2D object array suitable for spilling into Excel.
    /// </summary>
    private static object TensorToObjectArray(Tensor tensor)
    {
        var shape = tensor.shape;
        int rows;
        int cols;

        if (shape.Length == 1)
        {
            rows = 1;
            cols = (int)shape[0];
        }
        else if (shape.Length == 2)
        {
            rows = (int)shape[0];
            cols = (int)shape[1];
        }
        else
        {
            return $"#ERR: tensor rank {shape.Length} unsupported";
        }

        var data = tensor.data<float>().ToArray();
        var output = new object[rows, cols];
        int idx = 0;
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                output[r, c] = data[idx++];
            }
        }

        return output;
    }

    /// <summary>
    /// Formats weight and bias tensors into a readable matrix with headers for Excel.
    /// </summary>
    private static object BuildWeightMatrix(Tensor weight, Tensor bias)
    {
        if (ReferenceEquals(weight, null))
            return "#ERR: missing weight tensor";

        var shape = weight.shape;
        if (shape.Length != 2)
            return "#ERR: weight tensor must be 2D";

        int outDim = (int)shape[0];
        int inDim = (int)shape[1];
        bool hasBias = !ReferenceEquals(bias, null);
        int cols = inDim + 1 + (hasBias ? 1 : 0);
        var output = new object[outDim + 1, cols];

        output[0, 0] = "";
        for (int c = 0; c < inDim; c++)
        {
            output[0, c + 1] = $"in{c + 1}";
        }
        if (hasBias)
            output[0, inDim + 1] = "bias";

        var weightData = weight.data<float>().ToArray();
        float[] biasData = null;
        if (hasBias)
            biasData = bias.data<float>().ToArray();

        int idx = 0;
        for (int r = 0; r < outDim; r++)
        {
            output[r + 1, 0] = $"out{r + 1}";
            for (int c = 0; c < inDim; c++)
            {
                output[r + 1, c + 1] = weightData[idx++];
            }
            if (hasBias)
            {
                output[r + 1, inDim + 1] = biasData?[r];
            }
        }

        return output;
    }

    /// <summary>
    /// Normalizes Excel scalars and single-cell ranges to a plain object for comparison.
    /// </summary>
    private static object NormalizeScalar(object value)
    {
        if (value is ExcelReference xref)
        {
            var v = XlCall.Excel(XlCall.xlCoerce, xref);
            return NormalizeScalar(v);
        }

        if (value is object[,] arr && arr.GetLength(0) == 1 && arr.GetLength(1) == 1)
            return NormalizeScalar(arr[0, 0]);

        if (value is ExcelMissing || value is ExcelEmpty || value is null)
            return null;

        return value;
    }

    /// <summary>
    /// Normalizes a trigger value (including references) into a string token for change detection.
    /// </summary>
    private static string TriggerKey(object trigger)
    {
        if (trigger == null) return "<null>";
        if (trigger is ExcelMissing) return "<missing>";
        if (trigger is ExcelEmpty) return "<empty>";

        if (trigger is ExcelReference xref)
        {
            var v = XlCall.Excel(XlCall.xlCoerce, xref);
            return TriggerKey(v);
        }

        if (trigger is object[,] arr && arr.GetLength(0) == 1 && arr.GetLength(1) == 1)
            return TriggerKey(arr[0, 0]);

        return trigger.ToString();
    }

    [ExcelFunction(Name = "DL.TRIGGER_KEY", Description = "Debug: show normalized trigger key")]
    public static string TriggerKeyDebug(object trigger) => TriggerKey(trigger);

    /// <summary>
    /// Builds the default two-layer MLP used for quick-start scenarios.
    /// </summary>
    private static torch.nn.Module BuildDefaultMlp(int input, int hidden, int output)
    {
        return torch.nn.Sequential(
            ("fc1", torch.nn.Linear(input, hidden)),
            ("tanh1", torch.nn.Tanh()),
            ("fc2", torch.nn.Linear(hidden, output))
        );
    }

    /// <summary>
    /// Creates an optimizer for the model based on current settings.
    /// </summary>
    private static torch.optim.Optimizer CreateOptimizer(DlModelState model)
    {
        var optimizerName = model.OptimizerName ?? "adam";
        if (optimizerName.Equals("sgd", StringComparison.OrdinalIgnoreCase))
        {
            return torch.optim.SGD(model.TorchModel.parameters(), learningRate: model.LearningRate);
        }

        return torch.optim.Adam(model.TorchModel.parameters(), lr: model.LearningRate);
    }

    /// <summary>
    /// Adds a file to a zip archive with the specified entry name.
    /// </summary>
    private static void AddFileToArchive(ZipArchive archive, string sourcePath, string entryName)
    {
        var entry = archive.CreateEntry(entryName);
        using (var entryStream = entry.Open())
        using (var fileStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
        {
            fileStream.CopyTo(entryStream);
        }
    }

    /// <summary>
    /// Writes model metadata to disk as JSON.
    /// </summary>
    private static void WriteMetadata(string path, DlModelPersistence meta)
    {
        var serializer = new DataContractJsonSerializer(typeof(DlModelPersistence));
        using (var stream = new FileStream(path, FileMode.Create, FileAccess.Write))
        {
            serializer.WriteObject(stream, meta);
        }
    }

    /// <summary>
    /// Reads model metadata from a stream.
    /// </summary>
    private static DlModelPersistence ReadMetadata(Stream stream)
    {
        var serializer = new DataContractJsonSerializer(typeof(DlModelPersistence));
        return (DlModelPersistence)serializer.ReadObject(stream);
    }

    /// <summary>
    /// Saves a Torch state dictionary to disk using the available torch.save overload.
    /// </summary>
    private static void SaveStateDict(IDictionary<string, Tensor> stateDict, string path)
    {
        // TorchSharp 0.105 exposes torch.save for Tensor; use reflection to keep dictionary support if available.
        var saveMethod = typeof(torch).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .FirstOrDefault(m => m.Name == "save" && m.GetParameters().Length == 2);

        if (saveMethod != null)
        {
            saveMethod.Invoke(null, new object[] { stateDict, path });
        }
        else
        {
            throw new NotSupportedException("torch.save overload not found");
        }
    }

    /// <summary>
    /// Appends a log line to the add-in log file; failures are swallowed to keep UDFs safe.
    /// </summary>
    private static void Log(string message)
    {
        try
        {
            var line = $"{DateTime.Now:O} | {message}{Environment.NewLine}";
            File.AppendAllText(LogPathSafe, line, Encoding.UTF8);
        }
        catch { }
    }

    [ExcelFunction(Name = "DL.LOG_WRITE_TEST", Description = "Debug: force a log write")]
    public static string LogWriteTest()
    {
        try
        {
            Log("LOG_WRITE_TEST");
            return "attempted write to: " + LogPathSafe;
        }
        catch (Exception ex)
        {
            return "log write FAIL: " + ex.ToString();
        }
    }


    [ExcelFunction(Name = "DL.LOG_PATH", Description = "Debug: show where the log file should be written")]
    public static string LogPath() => LogPathSafe;

    [ExcelFunction(Name = "DL.TORCH_TEST", Description = "Debug: verify TorchSharp loads and can create a tensor")]
    public static string TorchTest()
    {
        try
        {
            EnsureTorch();
            using (var t = torch.ones(new long[] { 1 }))
            {
                return "torch ok: " + t.ToString();
            }
        }
        catch (Exception ex)
        {
            return "torch FAIL: " + ex.GetType().FullName + " | " + ex.Message;
        }
    }

    [ExcelFunction(Name = "DL.TORCH_TEST_DETAIL", Description = "Debug: TorchSharp init exception details")]
    public static string TorchTestDetail()
    {
        try
        {
            var baseDir = GetTorchBaseDir();
            var missing = GetMissingTorchNativeFiles(baseDir);
            if (missing.Count > 0)
            {
                return "torch native missing: " + string.Join(", ", missing) + " | dir=" + baseDir;
            }

            EnsureTorch();
            using (var t = torch.ones(new long[] { 1 }))
                return "torch ok: " + t.ToString();
        }
        catch (Exception ex)
        {
            return ex.ToString();
        }
    }

    [ExcelFunction(Name = "DL.TORCH_NATIVE_CHECK", Description = "Debug: list missing torch native DLLs")]
    public static string TorchNativeCheck()
    {
        var baseDir = GetTorchBaseDir();
        var missing = GetMissingTorchNativeFiles(baseDir);
        if (missing.Count == 0)
        {
            return "torch native ok: " + string.Join(", ", TorchNativeFiles) + " | dir=" + baseDir;
        }

        return "torch native missing: " + string.Join(", ", missing) + " | dir=" + baseDir;
    }

    [DataContract]
    private sealed class DlModelPersistence
    {
        [DataMember(Order = 1)]
        public int FormatVersion { get; set; }

        [DataMember(Order = 2)]
        public string ModelId { get; set; }

        [DataMember(Order = 3)]
        public string Description { get; set; }

        [DataMember(Order = 4)]
        public int InputDim { get; set; }

        [DataMember(Order = 5)]
        public int HiddenDim { get; set; }

        [DataMember(Order = 6)]
        public int OutputDim { get; set; }

        [DataMember(Order = 7)]
        public string OptimizerName { get; set; }

        [DataMember(Order = 8)]
        public double LearningRate { get; set; }

        [DataMember(Order = 9)]
        public string LastTriggerKey { get; set; }

        [DataMember(Order = 10)]
        public List<DlLossEntry> LossHistory { get; set; }
    }

    [DataContract]
    private sealed class DlLossEntry
    {
        [DataMember(Order = 1)]
        public int Epoch { get; set; }

        [DataMember(Order = 2)]
        public double Loss { get; set; }
    }

    [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr LoadLibrary(string lpFileName);
}
