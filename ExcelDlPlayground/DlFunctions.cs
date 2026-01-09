using ExcelDna.Integration;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using TorchSharp;
using static TorchSharp.torch;
using System.Runtime.InteropServices;

public static class DlFunctions
{
    private static bool _torchInitialized;

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

    private static void EnsureTorch()
    {
        if (_torchInitialized) return;

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

    [ExcelFunction(Name = "DL.MODEL_CREATE", Description = "Create a model and return a model_id")]
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
            var net = torch.nn.Sequential(
                ("fc1", torch.nn.Linear(input, hidden)),
                ("tanh1", torch.nn.Tanh()),
                ("fc2", torch.nn.Linear(hidden, output))
            );

            model.TorchModel = net;
            model.LossFn = torch.nn.BCEWithLogitsLoss(); // good default for XOR
            model.Optimizer = torch.optim.Adam(model.TorchModel.parameters(), lr: 0.1);

            Log($"ModelCreate | desc={description ?? "<null>"} | id={id} | in={input} hidden={hidden} out={output}");
            return id;
        }
        catch (Exception ex)
        {
            Log($"ModelCreate error | desc={description ?? "<null>"} | ex={ex}");
            return "#ERR: " + ex.Message;
        }
    }


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

        var functionName = nameof(Train);
        var parameters = new object[] { model_id, opts ?? "", key };

        return ExcelAsyncUtil.RunTask(functionName, parameters, async () =>
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
                EnsureTorch();

                if (model.TorchModel == null)
                {
                    return "#ERR: model not initialized. Call DL.MODEL_CREATE first.";
                }

                if (model.LossFn == null)
                {
                    model.LossFn = torch.nn.BCEWithLogitsLoss();
                }

                if (model.Optimizer == null)
                {
                    model.Optimizer = torch.optim.Adam(model.TorchModel.parameters(), lr: learningRate);
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
                        using (var output = model.TorchModel.forward(xTensor))
                        using (var lossTensor = model.LossFn.forward(output, yTensor))
                        {
                            lossTensor.backward();
                            model.Optimizer.step();
                            loss = lossTensor.ToSingle();
                        }

                        model.LossHistory.Add((e, loss));
                        await Task.Delay(1).ConfigureAwait(false);
                    }

                    model.UpdateWeightSnapshot();
                }

                model.LastTriggerKey = key;
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
                model.TrainLock.Release();
            }
        });
    }

    [ExcelFunction(Name = "DL.LOSS_HISTORY", Description = "Spill epoch/loss history for a model")]
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
            EnsureTorch();
            using (var t = torch.ones(new long[] { 1 }))
                return "torch ok: " + t.ToString();
        }
        catch (Exception ex)
        {
            return ex.ToString();
        }
    }

    [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr LoadLibrary(string lpFileName);
}
