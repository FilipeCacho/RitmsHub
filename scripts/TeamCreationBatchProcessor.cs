using System.Collections.Generic;
using System.Threading.Tasks;
using System.Threading;
using System;
using System.Linq;

public class TeamCreationBatchProcessor<T, TResult>
{
    private const int BatchSize = 10;
    private const int MaxRetries = 3;
    private const int RetryDelay = 5000; // 5 seconds
    private const int OperationTimeout = 300000; // 5 minutes
    private const int DelayBetweenOperations = 2000; // 2 seconds

    public async Task<List<TResult>> ProcessInBatchesAsync(List<T> items, Func<T, Task<TResult>> processFunc, string operationName)
    {
        int totalBatches = (int)Math.Ceiling((double)items.Count / BatchSize);
        int currentBatch = 0;
        List<TResult> results = new List<TResult>();

        for (int i = 0; i < items.Count; i += BatchSize)
        {
            var batch = items.Skip(i).Take(BatchSize).ToList();

            // Process the batch here
            currentBatch++;
            Console.WriteLine($"Processing {operationName} batch {currentBatch}/{totalBatches}...");

            var batchResults = await ProcessBatchWithRetriesAsync(batch, processFunc, operationName);
            results.AddRange(batchResults);

            Console.WriteLine($"Completed {operationName} batch {currentBatch}/{totalBatches}");

            // Delay between batches to avoid overloading the system
            await Task.Delay(DelayBetweenOperations);
        }

        return results;
    }

    private async Task<List<TResult>> ProcessBatchWithRetriesAsync(List<T> batch, Func<T, Task<TResult>> processFunc, string operationName)
    {
        List<TResult> batchResults = new List<TResult>();

        for (int attempt = 1; attempt <= MaxRetries; attempt++)
        {
            try
            {
                using (var cts = new CancellationTokenSource(OperationTimeout))
                {
                    foreach (var item in batch)
                    {
                        TResult result = await processFunc(item);
                        batchResults.Add(result);
                        await Task.Delay(DelayBetweenOperations, cts.Token);
                    }
                }
                return batchResults;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Attempt {attempt} for {operationName} batch failed: {ex.Message}");
                if (attempt < MaxRetries)
                {
                    Console.WriteLine($"Retrying in {RetryDelay / 1000} seconds...");
                    await Task.Delay(RetryDelay);
                }
            }
        }

        return batchResults;
    }
}