<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Schedule;

use DateTimeInterface;
use Mnb\PHPExcel\MnbExcel;

final class SpreadsheetScheduler
{
    public function __construct(private readonly SchedulerStoreInterface $scheduler)
    {
    }

    /** @param array<string,mixed>|string|null $database @param array<string,mixed> $options */
    public function scheduleImport(string $id, string $cron, string $path, array|string|null $database, string $table, array $options = []): void
    {
        $this->scheduler->add($id, $cron, 'spreadsheet.import', compact('path', 'database', 'table', 'options'));
    }

    /** @param list<array<int|string,mixed>> $rows @param array<string,mixed> $options */
    public function scheduleExport(string $id, string $cron, array $rows, string $path, array $options = []): void
    {
        $this->scheduler->add($id, $cron, 'spreadsheet.export', compact('rows', 'path', 'options'));
    }

    /** @return array<string,mixed> */
    public function runDue(?DateTimeInterface $now = null): array
    {
        return $this->scheduler->runDue(static function (string $type, array $payload): array {
            if ($type === 'spreadsheet.import') {
                return MnbExcel::largeImportToSql((string) $payload['path'], $payload['database'] ?? null, (string) $payload['table'], (array) ($payload['options'] ?? []));
            }
            if ($type === 'spreadsheet.export') {
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
            }
            throw new \RuntimeException('Unsupported scheduled spreadsheet task: ' . $type);
        }, $now);
    }


    /** Run due jobs continuously with locking, signal handling and bounded cycles. @param array<string,mixed> $options */
    public function runForever(array $options = []): array
    {
        return (new SchedulerRunner($this))->runForever($options);
    }

    /** @return array<string,mixed> */
    public function tasks(): array
    {
        return $this->scheduler->all();
    }
}
