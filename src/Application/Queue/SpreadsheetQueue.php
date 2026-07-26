<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Queue;

use Mnb\PHPExcel\MnbExcel;

final class SpreadsheetQueue
{
    public function __construct(private readonly FileQueue $queue)
    {
    }

    /** @param array<string,mixed>|string|null $database @param array<string,mixed> $options */
    public function enqueueImport(string $path, array|string|null $database, string $table, array $options = [], int $delaySeconds = 0): QueueJob
    {
        return $this->queue->enqueue('spreadsheet.import', compact('path', 'database', 'table', 'options'), $delaySeconds);
    }

    /** @param array<string,mixed>|string|null $database @param array<string,mixed> $options */
    public function enqueueDomainImport(string $domain, string $path, array|string|null $database = null, string $table = '', array $options = [], int $delaySeconds = 0): QueueJob
    {
        return $this->queue->enqueue('spreadsheet.domain_import', compact('domain', 'path', 'database', 'table', 'options'), $delaySeconds);
    }

    /** @param list<array<int|string,mixed>> $rows @param array<string,mixed> $options */
    public function enqueueExport(array $rows, string $path, array $options = [], int $delaySeconds = 0): QueueJob
    {
        return $this->queue->enqueue('spreadsheet.export', compact('rows', 'path', 'options'), $delaySeconds);
    }

    /** @param array<string,mixed> $options */
    public function work(array $options = []): array
    {
        $worker = new QueueWorker($this->queue);
        $worker->register('spreadsheet.import', static fn (array $payload): array => MnbExcel::largeImportToSql(
            (string) $payload['path'],
            $payload['database'] ?? null,
            (string) $payload['table'],
            (array) ($payload['options'] ?? [])
        ));
        $worker->register('spreadsheet.domain_import', static fn (array $payload): array => MnbExcel::importDomain(
            (string) $payload['domain'],
            (string) $payload['path'],
            $payload['database'] ?? null,
            (string) ($payload['table'] ?? ''),
            (array) ($payload['options'] ?? [])
        ));
        $worker->register('spreadsheet.export', static function (array $payload): array {
            $builder = MnbExcel::fromArray(array_values((array) ($payload['rows'] ?? [])));
            $options = (array) ($payload['options'] ?? []);
            if ((bool) ($options['with_header'] ?? true)) {
                $builder->withHeader();
            }
            if ((bool) ($options['auto_width'] ?? true)) {
                $builder->autoWidth();
            }
            $path = (string) $payload['path'];
            $builder->save($path);
            return ['path' => $path, 'size_bytes' => filesize($path) ?: 0];
        });
        return $worker->work($options);
    }
}
