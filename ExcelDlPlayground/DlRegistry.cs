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
    public readonly List<Tensor> WeightSnapshot = new List<Tensor>();
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
        foreach (var t in WeightSnapshot)
        {
            t.Dispose();
        }
        WeightSnapshot.Clear();

        if (TorchModel == null)
            return;

        foreach (var p in TorchModel.parameters())
        {
            WeightSnapshot.Add(p.detach().clone().cpu());
        }
    }
}
