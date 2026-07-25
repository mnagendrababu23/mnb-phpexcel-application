<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

final class ImportStatusReader
{
    /** @return array<string,mixed> */
    public static function read(string $manifestPath, array $options = []): array
    {
        if (!is_file($manifestPath)) {
            return [
                'status' => 'missing',
                'manifest_path' => $manifestPath,
                'exists' => false,
                'percent' => 0.0,
                'resume_ready' => false,
            ];
        }

        $json = file_get_contents($manifestPath);
        $data = is_string($json) ? json_decode($json, true) : null;
        if (!is_array($data)) {
            return [
                'status' => 'invalid',
                'manifest_path' => $manifestPath,
                'exists' => true,
                'percent' => 0.0,
                'resume_ready' => false,
            ];
        }

        $totalRows = (int) ($options['total_rows'] ?? $data['total_rows'] ?? $data['estimated_total_rows'] ?? 0);
        $rowsScanned = (int) ($data['rows_scanned'] ?? 0);
        $percent = $totalRows > 0 ? min(100.0, round(($rowsScanned / $totalRows) * 100, 2)) : null;
        $status = (string) ($data['status'] ?? 'unknown');

        return array_merge($data, [
            'exists' => true,
            'manifest_path' => $manifestPath,
            'status' => $status,
            'percent' => $percent,
            'resume_ready' => in_array($status, ['paused', 'failed', 'running'], true),
            'failed_rows_download_ready' => is_file((string) ($data['failed_rows_csv'] ?? '')),
        ]);
    }
}
