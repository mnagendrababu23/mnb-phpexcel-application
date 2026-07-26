<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Support\MnbExcelException;
use PDO;
use Throwable;

final class SpreadsheetApi
{
    /** @param array<string,mixed> $request @param PDO|array<string,mixed>|string|null $pdo @return array<string,mixed> */
    public function handle(string $action, array $request, PDO|array|string|null $pdo = null): array
    {
        try {
            $data = match (strtolower(trim($action))) {
                'upload' => (new AjaxUploadHandler())->handle($request['file'] ?? $request, (array) ($request['options'] ?? [])),
                'preview' => MnbExcel::previewDomainImport(
                    (string) ($request['domain'] ?? 'users'),
                    (string) ($request['path'] ?? ''),
                    (array) ($request['options'] ?? [])
                ),
                'import' => MnbExcel::importDomain(
                    (string) ($request['domain'] ?? 'users'),
                    (string) ($request['path'] ?? ''),
                    $pdo ?? ($request['database'] ?? null),
                    (string) ($request['table'] ?? ''),
                    (array) ($request['options'] ?? [])
                ),
                'import-many' => (new MultiFileImportManager())->importDomain(
                    (string) ($request['domain'] ?? 'users'),
                    array_values((array) ($request['files'] ?? [])),
                    $pdo ?? ($request['database'] ?? null),
                    (string) ($request['table'] ?? ''),
                    (array) ($request['options'] ?? [])
                ),
                'status' => MnbExcel::importDashboard((string) ($request['manifest'] ?? ''), (array) ($request['options'] ?? [])),
                'export' => $this->export($request),
                default => throw new MnbExcelException('Unknown spreadsheet API action: ' . $action),
            };

            return ['ok' => true, 'status' => 'success', 'http_status' => 200, 'data' => $data];
        } catch (Throwable $exception) {
            return [
                'ok' => false,
                'status' => 'error',
                'http_status' => $exception instanceof MnbExcelException ? 422 : 500,
                'error' => $exception->getMessage(),
                'exception' => $exception::class,
            ];
        }
    }

    /** @param array<string,mixed> $request @return array<string,mixed> */
    private function export(array $request): array
    {
        $path = (string) ($request['path'] ?? '');
        if ($path === '') {
            throw new MnbExcelException('Export path is required.');
        }
        $rows = array_values((array) ($request['rows'] ?? []));
        $builder = MnbExcel::fromArray($rows);
        if ((bool) ($request['with_header'] ?? true)) {
            $builder->withHeader();
        }
        if ((bool) ($request['auto_width'] ?? true)) {
            $builder->autoWidth();
        }
        $builder->save($path);
        return ['path' => $path, 'size_bytes' => filesize($path) ?: 0, 'sha256' => hash_file('sha256', $path)];
    }
}
