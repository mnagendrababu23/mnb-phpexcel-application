<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

use Mnb\PHPExcel\Large\LargeXlsxWriteSession;

final class BenchmarkSuite
{
    /** @return array<string,mixed> */
    public static function plan(array $options = []): array
    {
        $rows = array_values(array_map('intval', $options['rows'] ?? [100000, 500000, 1000000]));
        $columns = (int) ($options['columns'] ?? 10);

        return [
            'status' => 'ready',
            'note' => 'This is a reproducible benchmark plan. External library results are intentionally not hard-coded; run locally with the listed packages installed.',
            'datasets' => array_map(static fn (int $rowCount): array => [
                'rows' => $rowCount,
                'columns' => $columns,
                'shape' => $rowCount . 'x' . $columns,
                'recommended_mode' => $rowCount >= 500000 ? 'cli' : 'cli_or_long_running_worker',
            ], $rows),
            'libraries' => [
                'mnb_phpexcel_large_writer' => [
                    'package' => 'mnb/mnb-phpexcel',
                    'mode' => 'generator_to_large_xlsx',
                    'command' => 'php tools/benchmark-large-writer.php --rows={rows} --cols={columns}',
                    'measures' => ['elapsed_seconds', 'peak_memory_mb', 'output_size_mb', 'rows_per_second'],
                ],
                'mnb_phpexcel_large_import' => [
                    'package' => 'mnb/mnb-phpexcel',
                    'mode' => 'large_xlsx_to_streaming_chunks',
                    'command' => 'php tools/benchmark-large-reader.php --file=generated-{rows}.xlsx --chunk=5000',
                    'measures' => ['elapsed_seconds', 'peak_memory_mb', 'rows_per_second'],
                ],
                'phpoffice_phpspreadsheet' => [
                    'package' => 'phpoffice/phpspreadsheet',
                    'mode' => 'comparison_optional_dependency',
                    'install' => 'composer require --dev phpoffice/phpspreadsheet',
                    'note' => 'Use only when the optional comparison dependency is installed in a benchmark workspace.',
                ],
                'openspout_openspout' => [
                    'package' => 'openspout/openspout',
                    'mode' => 'comparison_optional_dependency',
                    'install' => 'composer require --dev openspout/openspout',
                    'note' => 'Use only when the optional comparison dependency is installed in a benchmark workspace.',
                ],
                'mk_j_php_xlsxwriter' => [
                    'package' => 'mk-j/php_xlsxwriter',
                    'mode' => 'comparison_optional_dependency',
                    'install' => 'composer require --dev mk-j/php_xlsxwriter',
                    'note' => 'Use only when the optional comparison dependency is installed in a benchmark workspace.',
                ],
                'rap2hpoutre_fast_excel' => [
                    'package' => 'rap2hpoutre/fast-excel',
                    'mode' => 'laravel_or_collection_comparison_optional_dependency',
                    'install' => 'composer require --dev rap2hpoutre/fast-excel',
                    'note' => 'Optional Laravel-friendly comparison; run in a compatible workspace.',
                ],
            ],
            'report_columns' => ['library', 'dataset', 'operation', 'elapsed_seconds', 'peak_memory_mb', 'output_size_mb', 'rows_per_second', 'status', 'notes'],
            'rules' => [
                'Do not publish benchmark numbers generated on different machines as direct comparisons.',
                'Run each benchmark at least three times and report median values.',
                'Use CLI, not a browser request, for 500k and 1M row benchmarks.',
                'Keep generated output files outside the release ZIP.',
            ],
        ];
    }

    /**
     * Run an internal MNB large writer benchmark. This intentionally benchmarks only
     * the local package; third-party libraries are optional external comparisons.
     *
     * @return array<string,mixed>
     */
    public static function runLargeWriter(int $rows, int $columns = 10, array $options = []): array
    {
        if ($rows < 1 || $columns < 1) {
            throw MnbExcelException::withCode('Benchmark rows and columns must be positive integers.', ErrorCode::VALIDATION_FAILED);
        }

        $path = (string) ($options['path'] ?? sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'mnb-benchmark-' . $rows . 'x' . $columns . '.xlsx');
        $progressEvery = max(1, (int) ($options['progress_every'] ?? 0));
        $progressEvents = 0;
        $start = microtime(true);
        $startMemory = memory_get_usage(true);

        $session = new LargeXlsxWriteSession(self::generateRows($rows, $columns), null, [
            'with_header' => true,
            'auto_split_sheets' => true,
            'strict_integrity_validation' => (bool) ($options['strict_integrity_validation'] ?? false),
        ]);

        if ($progressEvery > 0) {
            $session->progress(static function () use (&$progressEvents): void {
                $progressEvents++;
            }, $progressEvery);
        }

        $result = $session->save($path);
        $elapsed = microtime(true) - $start;
        $peak = memory_get_peak_usage(true);

        return [
            'status' => 'completed',
            'operation' => 'large_writer',
            'rows' => $rows,
            'columns' => $columns,
            'path' => $path,
            'elapsed_seconds' => round($elapsed, 4),
            'peak_memory_mb' => round($peak / 1048576, 3),
            'memory_delta_mb' => round(max(0, $peak - $startMemory) / 1048576, 3),
            'output_size_mb' => is_file($path) ? round(filesize($path) / 1048576, 3) : null,
            'rows_per_second' => $elapsed > 0 ? round($rows / $elapsed, 2) : null,
            'progress_events' => $progressEvents,
            'writer_result' => $result,
        ];
    }

    /** @return \Generator<int,array<string,mixed>> */
    private static function generateRows(int $rows, int $columns): \Generator
    {
        for ($i = 1; $i <= $rows; $i++) {
            $row = [];
            for ($c = 1; $c <= $columns; $c++) {
                $key = 'col_' . $c;
                $row[$key] = match ($c % 5) {
                    0 => 'Text ' . $i . '-' . $c,
                    1 => $i,
                    2 => round(($i * $c) / 7, 6),
                    3 => '2026-07-' . str_pad((string) (($i % 28) + 1), 2, '0', STR_PAD_LEFT),
                    default => 'ID-' . str_pad((string) $i, 8, '0', STR_PAD_LEFT),
                };
            }
            yield $row;
        }
    }
}
