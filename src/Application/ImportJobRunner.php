<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use PDO;

final class ImportJobRunner
{
    /** @param PDO|array<string,mixed>|string|null $pdo @param array<string,mixed> $options @return array<string,mixed> */
    public static function resume(string $manifestPath, PDO|array|string|null $pdo = null, array $options = []): array
    {
        $status = ImportStatusReader::read($manifestPath);
        if (($status['exists'] ?? false) !== true) {
            throw MnbExcelException::withCode('Import manifest not found: ' . $manifestPath, ErrorCode::FILE_NOT_FOUND);
        }
        $path = (string) ($status['source_path'] ?? '');
        $table = (string) ($status['table'] ?? '');
        if ($path === '' || $table === '') {
            throw MnbExcelException::withCode('Import manifest is missing source_path or table.', ErrorCode::VALIDATION_FAILED);
        }
        $options = array_merge($status, $options, [
            'manifest_path' => $manifestPath,
            'resume' => true,
        ]);
        return MnbExcel::largeImportToSql($path, $pdo, $table, $options);
    }
}
