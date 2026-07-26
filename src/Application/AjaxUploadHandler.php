<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

use Mnb\PHPExcel\Storage\StorageManager;
use Mnb\PHPExcel\Support\MnbExcelException;

final class AjaxUploadHandler
{
    /** @param array<string,mixed>|string $file @param array<string,mixed> $options @return array<string,mixed> */
    public function handle(array|string $file, array $options = []): array
    {
        $validation = UploadSafetyValidator::validate($file, $options);
        if (($validation['valid'] ?? false) !== true) {
            return [
                'ok' => false,
                'status' => 'validation_failed',
                'http_status' => 422,
                'errors' => $validation['errors'] ?? [],
                'warnings' => $validation['warnings'] ?? [],
                'validation' => $validation,
            ];
        }

        $source = (string) ($validation['path'] ?? '');
        $originalName = (string) ($validation['name'] ?? basename($source));
        $directory = (string) ($options['directory'] ?? StorageManager::path('upload_path'));
        if (!is_dir($directory) && !mkdir($directory, 0775, true) && !is_dir($directory)) {
            throw new MnbExcelException('Unable to create upload directory: ' . $directory);
        }

        $extension = strtolower((string) ($validation['extension'] ?? pathinfo($originalName, PATHINFO_EXTENSION)));
        $base = preg_replace('/[^A-Za-z0-9._-]+/', '_', pathinfo($originalName, PATHINFO_FILENAME)) ?: 'upload';
        $filename = (string) ($options['filename'] ?? ($base . '-' . gmdate('Ymd-His') . '-' . bin2hex(random_bytes(4)) . ($extension !== '' ? '.' . $extension : '')));
        $target = rtrim($directory, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . basename($filename);

        $moved = false;
        if (is_uploaded_file($source)) {
            $moved = move_uploaded_file($source, $target);
        }
        if (!$moved) {
            $moved = @rename($source, $target);
        }
        if (!$moved) {
            $moved = copy($source, $target);
        }
        if (!$moved || !is_file($target)) {
            throw new MnbExcelException('Unable to store uploaded spreadsheet.');
        }

        $response = [
            'ok' => true,
            'status' => 'stored',
            'http_status' => 201,
            'file' => [
                'path' => $target,
                'name' => basename($target),
                'original_name' => $originalName,
                'extension' => $extension,
                'size_bytes' => filesize($target) ?: 0,
                'sha256' => hash_file('sha256', $target),
            ],
            'warnings' => $validation['warnings'] ?? [],
        ];

        if (is_callable($options['after_store'] ?? null)) {
            $response['result'] = $options['after_store']($target, $response);
        }

        return $response;
    }

    /** @param array<string,mixed> $response */
    public function json(array $response, int $flags = JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE): string
    {
        $json = json_encode($response, $flags);
        if ($json === false) {
            throw new MnbExcelException('Unable to encode upload response as JSON.');
        }
        return $json;
    }
}
