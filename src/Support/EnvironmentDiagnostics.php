<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

use Mnb\PHPExcel\Support\Xml\XmlReader;

use Mnb\PHPExcel\Support\Zip\ZipArchive;

final class EnvironmentDiagnostics
{
    /**
     * Return a developer-friendly capability report for the current PHP runtime.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function check(array $options = []): array
    {
        $extensions = [
            'json' => extension_loaded('json'),
            'zip' => ZipArchive::nativeAvailable(),
            'xmlreader' => XmlReader::nativeAvailable(),
            'zip_fallback' => true,
            'xmlreader_fallback' => true,
            'pdo' => extension_loaded('pdo'),
            'mbstring' => extension_loaded('mbstring'),
            'iconv' => extension_loaded('iconv'),
        ];

        $tempDir = (string) ($options['temp_dir'] ?? sys_get_temp_dir());
        $checks = [];
        $checks[] = self::checkRow('php_version', version_compare(PHP_VERSION, '8.1.0', '>='), 'PHP ' . PHP_VERSION . ' detected; PHP 8.1+ is required.');
        $checks[] = self::checkRow('ext_json', $extensions['json'], 'ext-json is required for JSON import/export.');
        $checks[] = self::checkRow('ext_zip', $extensions['zip'], 'ext-zip is recommended for high-performance XLSX/ODS package access; a pure-PHP fallback is active.', 'warning');
        $checks[] = self::checkRow('ext_xmlreader', $extensions['xmlreader'], 'ext-xmlreader is recommended for true streaming XML/XLSX/ODS reads; a pure-PHP compatibility parser is active.', 'warning');
        $checks[] = self::checkRow('ext_pdo', $extensions['pdo'], 'ext-pdo is required only for SQL import/export helpers.', 'warning');
        $checks[] = self::checkRow('ext_mbstring_or_iconv', $extensions['mbstring'] || $extensions['iconv'], 'ext-mbstring or ext-iconv is recommended for CSV encoding conversion.', 'warning');
        $checks[] = self::checkRow('temp_dir_writable', is_dir($tempDir) && is_writable($tempDir), 'Temporary directory must be writable for atomic saves: ' . $tempDir);

        $failed = count(array_filter($checks, static fn (array $row): bool => $row['status'] === 'fail'));
        $warnings = count(array_filter($checks, static fn (array $row): bool => $row['status'] === 'warning'));
        $passed = count(array_filter($checks, static fn (array $row): bool => $row['status'] === 'pass'));

        $xlsxWriteReady = true;
        $xlsxReadReady = true;
        $sqlReady = $extensions['pdo'];
        $encodingReady = $extensions['mbstring'] || $extensions['iconv'];

        return [
            'status' => $failed > 0 ? 'fail' : ($warnings > 0 ? 'warning' : 'pass'),
            'php' => PHP_VERSION,
            'os' => PHP_OS_FAMILY,
            'sapi' => PHP_SAPI,
            'memory_limit' => ini_get('memory_limit') ?: '',
            'temp_dir' => $tempDir,
            'extensions' => $extensions,
            'capabilities' => [
                'csv_ready' => true,
                'json_ready' => $extensions['json'],
                'xml_ready' => true,
                'xlsx_write_ready' => $xlsxWriteReady,
                'xlsx_read_ready' => $xlsxReadReady,
                'xlsx_integrity_validation_ready' => true,
                'strict_xml_validation_ready' => true,
                'native_xlsx_streaming_ready' => $extensions['zip'] && $extensions['xmlreader'],
                'sql_helpers_ready' => $sqlReady,
                'encoding_conversion_ready' => $encodingReady,
                'comments_hyperlinks_writer_ready' => true,
            ],
            'checks' => $checks,
            'summary' => [
                'passed' => $passed,
                'warning' => $warnings,
                'failed' => $failed,
            ],
        ];
    }


    /**
     * Return a compact alert payload that applications can show before running XLSX/SQL workflows.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function alert(array $options = []): array
    {
        $report = self::check($options);
        $alerts = [];

        if (!($report['extensions']['zip'] ?? false)) {
            $alerts[] = [
                'level' => 'warning',
                'code' => 'MNB_EXT_ZIP_FALLBACK',
                'feature' => 'xlsx',
                'message' => 'ext-zip is missing. The pure-PHP ZIP fallback is active; enable ext-zip for lower memory use and faster very-large XLSX/ODS workloads.',
                'fix' => 'Optional: enable the PHP zip extension for native performance.',
            ];
        }

        if (!($report['extensions']['xmlreader'] ?? false)) {
            $alerts[] = [
                'level' => 'warning',
                'code' => 'MNB_EXT_XMLREADER_FALLBACK',
                'feature' => 'xlsx_read',
                'message' => 'ext-xmlreader is missing. The pure-PHP XML compatibility parser is active; enable XMLReader for true forward-only streaming on very large files.',
                'fix' => 'Optional: enable the PHP xmlreader extension for native streaming performance.',
            ];
        }

        $pdoSqlite = class_exists(\PDO::class) && in_array('sqlite', \PDO::getAvailableDrivers(), true);
        if (!$pdoSqlite) {
            $alerts[] = [
                'level' => 'warning',
                'code' => 'MNB_PDO_SQLITE_MISSING',
                'feature' => 'large_shared_strings',
                'message' => 'pdo_sqlite is missing. Large XLSX files with huge sharedStrings.xml cannot use the disk-backed shared-string cache and may be rejected to protect memory.',
                'fix' => 'Enable pdo_sqlite for safer very-large XLSX imports, especially on shared hosting and Windows/XAMPP.',
            ];
        }

        $status = 'pass';
        foreach ($alerts as $alert) {
            if ($alert['level'] === 'error') {
                $status = 'fail';
                break;
            }
            if ($status !== 'fail' && $alert['level'] === 'warning') {
                $status = 'warning';
            }
        }

        return [
            'status' => $status,
            'ready' => $status !== 'fail',
            'alerts' => $alerts,
            'missing' => [
                'ext_zip' => !($report['extensions']['zip'] ?? false),
                'ext_xmlreader' => !($report['extensions']['xmlreader'] ?? false),
                'pdo_sqlite' => !$pdoSqlite,
            ],
            'message' => self::alertMessageFromAlerts($alerts),
            'environment' => $report,
        ];
    }

    /** @param array<string,mixed> $options */
    public static function alertMessage(array $options = []): string
    {
        return (string) self::alert($options)['message'];
    }

    /** @param list<array<string,string>> $alerts */
    private static function alertMessageFromAlerts(array $alerts): string
    {
        if ($alerts === []) {
            return 'MNB PHPExcel environment check passed. XLSX and large import requirements are available.';
        }

        $lines = ['MNB PHPExcel environment alert:'];
        foreach ($alerts as $alert) {
            $lines[] = '- [' . strtoupper((string) $alert['level']) . '] ' . $alert['message'];
            if (($alert['fix'] ?? '') !== '') {
                $lines[] = '  Fix: ' . $alert['fix'];
            }
        }

        return implode(PHP_EOL, $lines);
    }

    /** @return array{name:string,status:string,message:string} */
    private static function checkRow(string $name, bool $ok, string $message, string $softStatus = 'fail'): array
    {
        return [
            'name' => $name,
            'status' => $ok ? 'pass' : $softStatus,
            'message' => $message,
        ];
    }
}
