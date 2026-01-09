using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using TorchSharp;
using static TorchSharp.torch;
using TorchSharp.Modules;

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
    public string LastTriggerKey { get; set; }  // <-- NEW

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

    public DlModelState(string description)
    {
        Description = description;
        LastTriggerKey = null;  // <-- NEW
    }

    public void UpdateWeightSnapshot()
    {
        ClearLayerSnapshots(WeightSnapshot);

        if (TorchModel == null)
            return;

        foreach (var layer in TorchModel.named_children())
        {
            if (layer.module is torch.nn.Linear linear)
            {
                var weight = linear.weight.detach().clone().cpu();
                Tensor bias = null;
                if (linear.bias is not null)
                {
                    bias = linear.bias.detach().clone().cpu();
                }
                WeightSnapshot[layer.name] = new LayerTensorSnapshot(weight, bias);
            }
        }
    }

    public void UpdateGradSnapshot()
    {
        ClearLayerSnapshots(GradSnapshot);

        if (TorchModel == null)
            return;

        foreach (var layer in TorchModel.named_children())
        {
            if (layer.module is torch.nn.Linear linear)
            {
                var weightGrad = linear.weight.grad();
                if (weightGrad is null)
                    continue;

                var weight = weightGrad.detach().clone().cpu();
                Tensor bias = null;
                var biasGrad = linear.bias?.grad();
                if (biasGrad is not null)
                {
                    bias = biasGrad.detach().clone().cpu();
                }

                GradSnapshot[layer.name] = new LayerTensorSnapshot(weight, bias);
            }
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
