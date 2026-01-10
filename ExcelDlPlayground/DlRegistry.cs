using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using TorchSharp;
using static TorchSharp.torch;
using TorchSharp.Modules;
using System.Reflection;

/// <summary>
/// In-memory registry managing model states keyed by generated identifiers.
/// </summary>
internal static class DlRegistry
{
    private static readonly ConcurrentDictionary<string, DlModelState> _models =
        new ConcurrentDictionary<string, DlModelState>();

    /// <summary>
    /// Creates a new model entry with a generated ID and initializes default state.
    /// </summary>
    /// <param name="description">Optional description used for the model state.</param>
    /// <returns>Newly generated model identifier.</returns>
    public static string CreateModel(string description)
    {
        var id = Guid.NewGuid().ToString("N");
        var state = new DlModelState(description);
        _models[id] = state;
        return id;
    }

    /// <summary>
    /// Attempts to retrieve a model state by identifier.
    /// </summary>
    /// <param name="modelId">Identifier to search.</param>
    /// <param name="state">Resolved state when found.</param>
    /// <returns>True when a model is present.</returns>
    public static bool TryGet(string modelId, out DlModelState state) =>
        _models.TryGetValue(modelId, out state);

    /// <summary>
    /// Inserts or replaces a model state for the given identifier.
    /// </summary>
    /// <param name="modelId">Identifier to set.</param>
    /// <param name="state">State to store.</param>
    public static void Upsert(string modelId, DlModelState state)
    {
        _models[modelId] = state;
    }
}

/// <summary>
/// Holds the runtime state and snapshots for a single TorchSharp model instance.
/// </summary>
internal sealed class DlModelState
{
    /// <summary>Free-form description provided when the model was created.</summary>
    public string Description { get; }

    /// <summary>Trigger token to avoid re-training when unchanged.</summary>
    public string LastTriggerKey { get; set; }

    /// <summary>Training lock to avoid concurrent training on the same model.</summary>
    public readonly SemaphoreSlim TrainLock = new SemaphoreSlim(1, 1);

    /// <summary>Loss history for DL.LOSS_HISTORY.</summary>
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

    /// <summary>
    /// Initializes tracking state for a new model description.
    /// </summary>
    /// <param name="description">Free-form description provided by the user.</param>
    public DlModelState(string description)
    {
        Description = description;
        LastTriggerKey = null;
        IsTraining = false;
        TrainingVersion = 0;
        LastEpoch = 0;
        LastLoss = double.NaN;
    }

    /// <summary>
    /// Copies current layer weights into a CPU snapshot for later inspection in Excel.
    /// </summary>
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

    /// <summary>
    /// Copies current layer gradients into a CPU snapshot for inspection.
    /// </summary>
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

    /// <summary>
    /// Replaces cached activation tensors with values produced during a forward pass.
    /// </summary>
    /// <param name="activations">Layer-name to tensor mapping to cache.</param>
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

    /// <summary>
    /// Disposes and clears a snapshot dictionary to prevent native memory leaks.
    /// </summary>
    /// <param name="snapshots">Dictionary of layer snapshots to clear.</param>
    private static void ClearLayerSnapshots(Dictionary<string, LayerTensorSnapshot> snapshots)
    {
        foreach (var entry in snapshots)
        {
            entry.Value.Dispose();
        }
        snapshots.Clear();
    }

    /// <summary>
    /// Retrieves the gradient tensor from a TorchSharp parameter via reflection to stay version-tolerant.
    /// </summary>
    /// <param name="parameter">Parameter tensor whose gradient is requested.</param>
    /// <returns>Gradient tensor or null when absent.</returns>
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

/// <summary>
/// Lightweight disposable container for weight/bias snapshots used in Excel inspection UDFs.
/// </summary>
internal sealed class LayerTensorSnapshot : IDisposable
{
    public Tensor Weight { get; }
    public Tensor Bias { get; }

    /// <summary>
    /// Captures the provided tensors for later disposal.
    /// </summary>
    /// <param name="weight">Weight tensor snapshot.</param>
    /// <param name="bias">Bias tensor snapshot (optional).</param>
    public LayerTensorSnapshot(Tensor weight, Tensor bias)
    {
        Weight = weight;
        Bias = bias;
    }

    /// <summary>
    /// Disposes captured tensors to release native resources.
    /// </summary>
    public void Dispose()
    {
        Weight?.Dispose();
        Bias?.Dispose();
    }
}
