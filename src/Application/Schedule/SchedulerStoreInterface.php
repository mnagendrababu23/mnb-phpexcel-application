<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Schedule;

use DateTimeInterface;

interface SchedulerStoreInterface
{
    /** @param array<string,mixed> $payload */
    public function add(string $id, string $cron, string $type, array $payload, bool $enabled = true): void;
    public function remove(string $id): void;
    /** @param callable(string,array<string,mixed>,array<string,mixed>):mixed $dispatcher @return array<string,mixed> */
    public function runDue(callable $dispatcher, ?DateTimeInterface $now = null): array;
    /** @return array<string,mixed> */
    public function all(): array;
}
