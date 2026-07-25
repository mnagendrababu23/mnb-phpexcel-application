<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Storage;

use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

/**
 * Centralized storage paths for real applications. Keeps temp/manifests/failed uploads out of vendor.
 */
final class StorageManager
{
    /** @var array<string,string> */
    private static array $paths = [];

    /** @param array<string,string> $paths @return array<string,string> */
    public static function configure(array $paths = []): array
    {
        $base = (string) ($paths['base_path'] ?? self::$paths['base_path'] ?? (getcwd() ?: sys_get_temp_dir()) . DIRECTORY_SEPARATOR . 'mnb-phpexcel-storage');
        self::$paths = array_merge(self::defaults($base), self::$paths, self::normalizePaths($paths, $base));
        self::ensureDirectories();
        return self::$paths;
    }

    /** @return array<string,string> */
    public static function paths(): array
    {
        if (self::$paths === []) {
            self::configure();
        }
        return self::$paths;
    }

    public static function path(string $key, string $filename = ''): string
    {
        $paths = self::paths();
        if (!isset($paths[$key])) {
            throw MnbExcelException::withCode('Unknown storage path key: ' . $key, ErrorCode::FILE_WRITE_FAILED);
        }
        $base = rtrim($paths[$key], DIRECTORY_SEPARATOR);
        return $filename === '' ? $base : $base . DIRECTORY_SEPARATOR . self::safeRelativeName($filename);
    }

    public static function manifestPath(string $name): string
    {
        return self::path('manifest_path', self::safeBasename($name, 'import') . '.json');
    }

    public static function failedRowsPath(string $name): string
    {
        return self::path('failed_rows_path', self::safeBasename($name, 'failed-rows') . '.csv');
    }

    public static function tempPath(string $name): string
    {
        return self::path('temp_path', self::safeBasename($name, 'temp'));
    }

    public static function ensureDirectories(): void
    {
        foreach (self::$paths as $key => $path) {
            if (!str_ends_with($key, '_path')) {
                continue;
            }
            if (!is_dir($path) && !@mkdir($path, 0775, true) && !is_dir($path)) {
                throw MnbExcelException::withCode('Unable to create storage directory: ' . $path, ErrorCode::FILE_WRITE_FAILED);
            }
        }
    }

    /** @return array<string,string> */
    private static function defaults(string $base): array
    {
        return [
            'base_path' => $base,
            'temp_path' => $base . DIRECTORY_SEPARATOR . 'temp',
            'upload_path' => $base . DIRECTORY_SEPARATOR . 'uploads',
            'manifest_path' => $base . DIRECTORY_SEPARATOR . 'manifests',
            'failed_rows_path' => $base . DIRECTORY_SEPARATOR . 'failed-rows',
            'reports_path' => $base . DIRECTORY_SEPARATOR . 'reports',
        ];
    }

    /** @param array<string,string> $paths @return array<string,string> */
    private static function normalizePaths(array $paths, string $base): array
    {
        $normalized = [];
        foreach ($paths as $key => $path) {
            if (!is_string($path) || trim($path) === '') {
                continue;
            }
            $normalized[$key] = self::absolutePath($path, $base);
        }
        return $normalized;
    }

    private static function absolutePath(string $path, string $base): string
    {
        if (preg_match('/^[A-Za-z]:[\\\\\/]/', $path) === 1 || str_starts_with($path, DIRECTORY_SEPARATOR)) {
            return rtrim($path, DIRECTORY_SEPARATOR);
        }
        return rtrim($base, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . trim($path, DIRECTORY_SEPARATOR);
    }

    private static function safeRelativeName(string $name): string
    {
        $name = str_replace(['..', '/', '\\'], '_', trim($name));
        return $name === '' ? 'mnb-file' : $name;
    }

    private static function safeBasename(string $name, string $fallback): string
    {
        $base = pathinfo($name, PATHINFO_FILENAME) ?: $fallback;
        $base = preg_replace('/[^A-Za-z0-9_.-]+/', '-', $base) ?: $fallback;
        return trim($base, '.-') ?: $fallback;
    }
}
