<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Queue;

interface QueueBackendInterface
{
    /** @param array<string,mixed> $payload */
    public function enqueue(string $type, array $payload, int $delaySeconds = 0): QueueJob;

    public function reserve(int $visibilityTimeoutSeconds = 300, ?string $workerId = null): ?QueueJob;

    /** @param array<string,mixed> $result */
    public function complete(QueueJob $job, array $result = []): void;

    public function fail(QueueJob $job, \Throwable $exception, int $retryDelaySeconds = 0, int $maxAttempts = 3): void;

    public function releaseExpired(int $visibilityTimeoutSeconds = 300): int;

    /** @return array<string,int> */
    public function stats(): array;
}
