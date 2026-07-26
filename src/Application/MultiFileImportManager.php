<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

use Mnb\PHPExcel\Domain\DomainImportType;
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\MnbExcelException;
use PDO;
use Throwable;

final class MultiFileImportManager
{
    /**
     * @param list<string|array<string,mixed>> $files
     * @param callable(string|array<string,mixed>,int,array<string,mixed>):array<string,mixed> $importer
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public function run(array $files, callable $importer, array $options = []): array
    {
        if ($files === []) {
            throw new MnbExcelException('At least one import file is required.');
        }

        $stopOnError = (bool) ($options['stop_on_error'] ?? false);
        $validateUploads = (bool) ($options['validate_uploads'] ?? true);
        $progress = is_callable($options['progress'] ?? null) ? $options['progress'] : null;
        $results = [];
        $errors = [];
        $totals = [
            'files_total' => count($files),
            'files_completed' => 0,
            'files_failed' => 0,
            'rows_scanned' => 0,
            'rows_inserted' => 0,
            'rows_updated' => 0,
            'rows_skipped' => 0,
            'rows_failed' => 0,
        ];

        foreach (array_values($files) as $index => $file) {
            $path = is_array($file) ? (string) ($file['tmp_name'] ?? $file['path'] ?? '') : (string) $file;
            $name = is_array($file) ? (string) ($file['name'] ?? basename($path)) : basename($path);
            $state = ['index' => $index, 'number' => $index + 1, 'name' => $name, 'path' => $path];

            try {
                if ($validateUploads) {
                    $validation = UploadSafetyValidator::validate($file, array_replace(['allowed_extensions' => ['xlsx', 'xls', 'ods', 'csv', 'tsv', 'json', 'xml']], (array) ($options['upload_options'] ?? [])));
                    if (($validation['valid'] ?? false) !== true) {
                        throw new MnbExcelException('Import file failed validation: ' . implode('; ', (array) ($validation['errors'] ?? [])));
                    }
                    $state['validation'] = $validation;
                }

                $result = $importer($file, $index, $state);
                $results[] = $state + ['status' => 'completed', 'result' => $result];
                $totals['files_completed']++;
                foreach (['rows_scanned', 'rows_inserted', 'rows_updated', 'rows_skipped', 'rows_failed'] as $key) {
                    $totals[$key] += (int) ($result[$key] ?? $result['summary'][$key] ?? 0);
                }
            } catch (Throwable $exception) {
                $entry = $state + [
                    'status' => 'failed',
                    'error' => $exception->getMessage(),
                    'exception' => $exception::class,
                ];
                $results[] = $entry;
                $errors[] = $entry;
                $totals['files_failed']++;
                if ($stopOnError) {
                    break;
                }
            }

            if ($progress !== null) {
                $progress($totals + ['current_file' => $state, 'results' => $results]);
            }
        }

        return [
            'status' => $totals['files_failed'] === 0 ? 'completed' : ($totals['files_completed'] > 0 ? 'completed_with_errors' : 'failed'),
            'summary' => $totals,
            'results' => $results,
            'errors' => $errors,
        ];
    }

    /** @param list<string|array<string,mixed>> $files @param PDO|array<string,mixed>|string|null $pdo @param array<string,mixed> $options */
    public function importToSql(array $files, PDO|array|string|null $pdo, string $table, array $options = []): array
    {
        return $this->run($files, static function (string|array $file, int $index) use ($pdo, $table, $options): array {
            $path = is_array($file) ? (string) ($file['tmp_name'] ?? $file['path'] ?? '') : $file;
            $resolvedTable = is_callable($options['table_resolver'] ?? null)
                ? (string) $options['table_resolver']($path, $index, $table)
                : $table;
            return MnbExcel::largeImportToSql($path, $pdo, $resolvedTable, (array) ($options['import_options'] ?? $options));
        }, $options);
    }

    /** @param list<string|array<string,mixed>> $files @param PDO|array<string,mixed>|string|null $pdo @param array<string,mixed> $options */
    public function importDomain(DomainImportType|string $domain, array $files, PDO|array|string|null $pdo = null, string $table = '', array $options = []): array
    {
        return $this->run($files, static function (string|array $file, int $index) use ($domain, $pdo, $table, $options): array {
            $path = is_array($file) ? (string) ($file['tmp_name'] ?? $file['path'] ?? '') : $file;
            $resolvedTable = is_callable($options['table_resolver'] ?? null)
                ? (string) $options['table_resolver']($path, $index, $table)
                : $table;
            return MnbExcel::importDomain($domain, $path, $pdo, $resolvedTable, (array) ($options['import_options'] ?? $options));
        }, $options);
    }
}
