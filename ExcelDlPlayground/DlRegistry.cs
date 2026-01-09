using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using TorchSharp;
using static TorchSharp.torch;
using TorchSharp.Modules;
using System.Reflection;

internal static class DlRegistry
{
    private static readonly ConcurrentDictionary<string, DlModelState> _models =
        new ConcurrentDictionary<string, DlModelState>();

    public static string CreateModel(string description)
    {
        var id = Guid.NewGuid().ToString("N");
        var state = new DlModelState(description);
        _models[id] = state;
        return id;
    }

    public static bool TryGet(string modelId, out DlModelState state) =>
        _models.TryGetValue(modelId, out state);

    public static void Upsert(string modelId, DlModelState state)
    {
        _models[modelId] = state;
    }
}

internal sealed class DlModelState
{
    public string Description { get; }

    // Trigger token to avoid re-training
    public string LastTriggerKey { get; set; }

    // Training lock to avoid concurrent training on same model
    public readonly SemaphoreSlim TrainLock = new SemaphoreSlim(1, 1);

    // Loss history for DL.LOSS_HISTORY
    public readonly List<(int epoch, double loss)> LossHistory = new List<(int, double)>();
    public torch.nn.Module TorchModel { get; set; }
    public torch.optim.Optimizer Optimizer { get; set; }
    public torch.nn.Module LossFn { get; set; }
    public readonly Dictionary<string, LayerTensorSnapshot> WeightSnapshot =
        new Dictionary<string, LayerTensorSnapshot>(StringComparer.OrdinalIgnoreCase);
    public readonly Dictionary<string, LayerTensorSnapshot> GradSnapshot =
        new Dictionary<string, LayerTensorSnapshot>(StringComparer.OrdinalIgnoreCase);
    public readonly Dictionary<string, Tensor> ActivationSnapshot =
        new Dictionary<string, Tensor>(StringComparer.OrdinalIgnoreCase);
    public string OptimizerName { get; set; }
    public double LearningRate { get; set; }
    public int InputDim { get; set; }
    public int HiddenDim { get; set; }
    public int OutputDim { get; set; }

    // Training state
    public bool IsTraining { get; set; }
    public long TrainingVersion { get; set; }
    public int LastEpoch { get; set; }
    public double LastLoss { get; set; }

    public DlModelState(string description)
    {
        Description = description;
        LastTriggerKey = null;
        IsTraining = false;
        TrainingVersion = 0;
        LastEpoch = 0;
        LastLoss = double.NaN;
    }

    public void UpdateWeightSnapshot()
    {
        ClearLayerSnapshots(WeightSnapshot);

        if (TorchModel == null)
            return;

        foreach (var layer in TorchModel.named_children())
        {
            var linear = layer.module as Linear;
            if (linear == null)
                continue;

            var weight = linear.weight.detach().clone().cpu();
            Tensor bias = null;
            if (!ReferenceEquals(linear.bias, null))
            {
                bias = linear.bias.detach().clone().cpu();
            }
            WeightSnapshot[layer.name] = new LayerTensorSnapshot(weight, bias);
        }
    }

    public void UpdateGradSnapshot()
    {
        ClearLayerSnapshots(GradSnapshot);

        if (TorchModel == null)
            return;

        foreach (var layer in TorchModel.named_children())
        {
            var linear = layer.module as Linear;
            if (linear == null)
                continue;

            var weightGrad = GetGrad(linear.weight);
            if (ReferenceEquals(weightGrad, null))
                continue;

            var weight = weightGrad.detach().clone().cpu();
            Tensor bias = null;
            var biasGrad = GetGrad(linear.bias);
            if (!ReferenceEquals(biasGrad, null))
            {
                bias = biasGrad.detach().clone().cpu();
            }

            GradSnapshot[layer.name] = new LayerTensorSnapshot(weight, bias);
        }
    }

    public void UpdateActivationSnapshot(Dictionary<string, Tensor> activations)
    {
        foreach (var entry in ActivationSnapshot)
        {
            entry.Value.Dispose();
        }
        ActivationSnapshot.Clear();

        foreach (var entry in activations)
        {
            ActivationSnapshot[entry.Key] = entry.Value;
        }
    }

    private static void ClearLayerSnapshots(Dictionary<string, LayerTensorSnapshot> snapshots)
    {
        foreach (var entry in snapshots)
        {
            entry.Value.Dispose();
        }
        snapshots.Clear();
    }

    private static Tensor GetGrad(Tensor parameter)
    {
        if (ReferenceEquals(parameter, null))
            return null;

        var type = parameter.GetType();
        var prop = type.GetProperty("grad");
        if (prop != null)
            return prop.GetValue(parameter) as Tensor;

        var method = type.GetMethod("grad", Type.EmptyTypes);
        if (method != null)
            return method.Invoke(parameter, null) as Tensor;

        return null;
    }
}

internal sealed class LayerTensorSnapshot : IDisposable
{
    public Tensor Weight { get; }
    public Tensor Bias { get; }

    public LayerTensorSnapshot(Tensor weight, Tensor bias)
    {
        Weight = weight;
        Bias = bias;
    }

    public void Dispose()
    {
        Weight?.Dispose();
        Bias?.Dispose();
    }
}
