<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

/**
 * Converts a resumable import manifest into an API/admin-dashboard friendly shape.
 */
final class ImportDashboardHelper
{
    /** @param array<string,mixed> $options @return array<string,mixed> */
    public static function response(string|array $manifestOrStatus, array $options = []): array
    {
        $status = is_array($manifestOrStatus) ? $manifestOrStatus : ImportStatusReader::read($manifestOrStatus, $options);
        $state = (string) ($status['status'] ?? 'unknown');
        $rowsScanned = (int) ($status['rows_scanned'] ?? 0);
        $inserted = (int) ($status['inserted_rows'] ?? 0);
        $failed = (int) ($status['failed_rows'] ?? 0);
        $total = (int) ($options['total_rows'] ?? $status['total_rows'] ?? $status['estimated_total_rows'] ?? 0);
        $percent = $status['percent'] ?? ($total > 0 ? min(100.0, round(($rowsScanned / $total) * 100, 2)) : null);
        $failedPath = (string) ($status['failed_rows_csv'] ?? '');
        $manifestPath = (string) ($status['manifest_path'] ?? '');
        $baseDownloadUrl = rtrim((string) ($options['download_base_url'] ?? ''), '/');
        $failedDownloadUrl = null;
        if ($failedPath !== '' && is_file($failedPath)) {
            $failedDownloadUrl = (string) ($options['failed_rows_download_url'] ?? '');
            if ($failedDownloadUrl === '' && $baseDownloadUrl !== '') {
                $failedDownloadUrl = $baseDownloadUrl . '/' . rawurlencode(basename($failedPath));
            }
        }

        return [
            'status' => $state,
            'ok' => in_array($state, ['running', 'paused', 'completed'], true),
            'message' => self::message($state, $status),
            'safe_admin_message' => self::safeAdminMessage($state, $status),
            'progress' => [
                'percent' => $percent,
                'rows_scanned' => $rowsScanned,
                'inserted_rows' => $inserted,
                'failed_rows' => $failed,
                'valid_rows' => (int) ($status['valid_rows'] ?? 0),
                'chunks_completed' => (int) ($status['chunks_completed'] ?? 0),
                'total_rows' => $total > 0 ? $total : null,
                'estimated_remaining_seconds' => self::etaSeconds($status, $rowsScanned, $total),
            ],
            'files' => [
                'manifest_path' => $manifestPath,
                'failed_rows_path' => $failedPath !== '' ? $failedPath : null,
                'failed_rows_exists' => $failedPath !== '' && is_file($failedPath),
                'failed_rows_download_url' => $failedDownloadUrl !== '' ? $failedDownloadUrl : null,
            ],
            'resume' => [
                'ready' => (bool) ($status['resume_ready'] ?? in_array($state, ['paused', 'failed', 'running'], true)),
                'manifest_path' => $manifestPath,
                'last_row_number' => (int) ($status['last_row_number'] ?? 0),
                'recommended_action' => self::resumeAction($state),
            ],
            'raw' => ($options['include_raw'] ?? false) ? $status : null,
        ];
    }

    /** @param array<string,mixed> $status */
    private static function message(string $state, array $status): string
    {
        return match ($state) {
            'completed' => 'Import completed successfully.',
            'paused' => 'Import paused safely and can be resumed.',
            'running' => 'Import is running.',
            'failed' => 'Import failed. Review the error and resume after fixing the issue.',
            'missing' => 'Import status manifest was not found.',
            'invalid' => 'Import status manifest is invalid.',
            default => 'Import status is unknown.',
        };
    }

    /** @param array<string,mixed> $status */
    private static function safeAdminMessage(string $state, array $status): string
    {
        if ($state === 'failed') {
            $reason = (string) ($status['last_stop_reason'] ?? 'import_error');
            return 'The import stopped with reason: ' . $reason . '. Check logs or failed-row CSV before retrying.';
        }
        if ($state === 'paused') {
            return 'The import paused before timeout. Click resume to continue from the last saved row.';
        }
        if ($state === 'completed') {
            return 'All available rows were processed. Download failed rows if any corrections are needed.';
        }
        return self::message($state, $status);
    }

    /** @param array<string,mixed> $status */
    private static function etaSeconds(array $status, int $rowsScanned, int $total): ?int
    {
        if ($rowsScanned <= 0 || $total <= 0 || $rowsScanned >= $total) {
            return null;
        }
        $elapsed = (float) ($status['elapsed_seconds'] ?? $status['last_elapsed_seconds'] ?? 0);
        if ($elapsed <= 0) {
            $created = isset($status['created_at']) ? strtotime((string) $status['created_at']) : false;
            $updated = isset($status['updated_at']) ? strtotime((string) $status['updated_at']) : time();
            if ($created !== false && $updated !== false && $updated > $created) {
                $elapsed = (float) ($updated - $created);
            }
        }
        if ($elapsed <= 0) {
            return null;
        }
        $rate = $rowsScanned / $elapsed;
        if ($rate <= 0) {
            return null;
        }
        return (int) ceil(($total - $rowsScanned) / $rate);
    }

    private static function resumeAction(string $state): string
    {
        return match ($state) {
            'paused' => 'resume_now',
            'failed' => 'fix_error_then_resume',
            'running' => 'refresh_status',
            default => 'none',
        };
    }
}
