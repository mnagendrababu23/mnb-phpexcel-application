<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Queue;

use Mnb\PHPExcel\Support\MnbExcelException;
use Throwable;

final class QueueWorker
{
    /** @var array<string,callable(array<string,mixed>,QueueJob):array<string,mixed>|mixed> */
    private array $handlers = [];

    public function __construct(private readonly FileQueue $queue)
    {
    }

    /** @param callable(array<string,mixed>,QueueJob):array<string,mixed>|mixed $handler */
    public function register(string $type, callable $handler): self
    {
        $this->handlers[$type] = $handler;
        return $this;
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function work(array $options = []): array
    {
        $maxJobs = max(1, (int) ($options['max_jobs'] ?? 100));
        $timeBudget = max(0.0, (float) ($options['time_budget_seconds'] ?? 0));
        $maxAttempts = max(1, (int) ($options['max_attempts'] ?? 3));
        $retryDelay = max(0, (int) ($options['retry_delay_seconds'] ?? 30));
        $started = microtime(true);
        $processed = $completed = $failed = 0;
        $results = [];

        while ($processed < $maxJobs && ($timeBudget <= 0 || microtime(true) - $started < $timeBudget)) {
            $job = $this->queue->reserve();
            if ($job === null) {
                break;
            }
            $processed++;
            try {
                $handler = $this->handlers[$job->type] ?? null;
                if ($handler === null) {
                    throw new MnbExcelException('No queue handler registered for job type: ' . $job->type);
                }
                $result = $handler($job->payload, $job);
                $normalized = is_array($result) ? $result : ['value' => $result];
                $this->queue->complete($job, $normalized);
                $completed++;
                $results[] = ['id' => $job->id, 'type' => $job->type, 'status' => 'completed', 'result' => $normalized];
            } catch (Throwable $exception) {
                $this->queue->fail($job, $exception, $retryDelay, $maxAttempts);
                $failed++;
                $results[] = ['id' => $job->id, 'type' => $job->type, 'status' => 'failed', 'error' => $exception->getMessage()];
            }
        }

        return [
            'status' => $failed === 0 ? 'completed' : 'completed_with_errors',
            'processed' => $processed,
            'completed' => $completed,
            'failed' => $failed,
            'elapsed_seconds' => round(microtime(true) - $started, 6),
            'queue' => $this->queue->stats(),
            'results' => $results,
        ];
    }
}
