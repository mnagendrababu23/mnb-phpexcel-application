<?php

declare(strict_types=1);

namespace Mnb\PHPExcel;

use Mnb\PHPExcel\Application\ImportJobRunner;
use Mnb\PHPExcel\Application\ImportProfile;
use Mnb\PHPExcel\Application\ImportProfileManager;
use Mnb\PHPExcel\Application\ImportStatusReader;
use Mnb\PHPExcel\Application\ImportDashboardHelper;
use Mnb\PHPExcel\Application\LoggerBridge;
use Mnb\PHPExcel\Application\RowTransformerPipeline;
use Mnb\PHPExcel\Application\UploadSafetyValidator;
use Mnb\PHPExcel\Application\AjaxUploadHandler;
use Mnb\PHPExcel\Application\AjaxUploader;
use Mnb\PHPExcel\Application\MultiFileImportManager;
use Mnb\PHPExcel\Application\SpreadsheetApi;
use Mnb\PHPExcel\Application\Mail\CallbackMailer;
use Mnb\PHPExcel\Application\Mail\MailerInterface;
use Mnb\PHPExcel\Application\Mail\NativeMailer;
use Mnb\PHPExcel\Application\Mail\SpreadsheetMailer;
use Mnb\PHPExcel\Application\Mail\SmtpMailer;
use Mnb\PHPExcel\Application\Queue\FileQueue;
use Mnb\PHPExcel\Application\Queue\PdoQueue;
use Mnb\PHPExcel\Application\Queue\QueueBackendInterface;
use Mnb\PHPExcel\Application\Queue\SpreadsheetQueue;
use Mnb\PHPExcel\Application\Schedule\FileScheduler;
use Mnb\PHPExcel\Application\Schedule\PdoScheduler;
use Mnb\PHPExcel\Application\Schedule\SchedulerRunner;
use Mnb\PHPExcel\Application\Schedule\SpreadsheetScheduler;
use Mnb\PHPExcel\Application\Http\HttpResponse;
use Mnb\PHPExcel\Application\Http\SpreadsheetHttpEndpoint;
use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookBuilder;
use Mnb\PHPExcel\Reader\CsvReader;
use Mnb\PHPExcel\Reader\JsonReader;
use Mnb\PHPExcel\Reader\ReadSession;
use Mnb\PHPExcel\Reader\ReaderRegistry;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\OdsReader;
use Mnb\PHPExcel\Compatibility\XlsReader;
use Mnb\PHPExcel\Contracts\ReaderPluginInterface;
use Mnb\PHPExcel\Reader\XmlReader;
use Mnb\PHPExcel\Reader\XlsxReader;
use Mnb\PHPExcel\Reader\XlsxInspector;
use Mnb\PHPExcel\Events\EventDispatcher;
use Mnb\PHPExcel\Domain\DomainImportPreset;
use Mnb\PHPExcel\Domain\DomainImportRegistry;
use Mnb\PHPExcel\Domain\DomainImportType;
use Mnb\PHPExcel\Import\DomainImporter;
use Mnb\PHPExcel\Import\ImportQualityAnalyzer;
use Mnb\PHPExcel\Large\ImportMethodAdvisor;
use Mnb\PHPExcel\Large\LargeExcelDatabaseImportEngine;
use Mnb\PHPExcel\Large\LargeImportManifest;
use Mnb\PHPExcel\Large\LargeExcelPreflightAnalyzer;
use Mnb\PHPExcel\Large\LargeXlsxReadSession;
use Mnb\PHPExcel\Large\LargeXlsxWriteSession;
use Mnb\PHPExcel\Large\LargePdoCursor;
use Mnb\PHPExcel\Plugin\MnbExcelPluginInterface;
use Mnb\PHPExcel\Plugin\PluginManager;
use Mnb\PHPExcel\Security\CellSafetyScanner;
use Mnb\PHPExcel\Support\DatabaseConfigResolver;
use Mnb\PHPExcel\Support\DatabaseConnectionFactory;
use Mnb\PHPExcel\Support\EncodingDetector;
use Mnb\PHPExcel\Support\EnvironmentDiagnostics;
use Mnb\PHPExcel\Support\FileFormatDetector;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\ErrorReporter;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\ReleaseReadiness;
use Mnb\PHPExcel\Support\SheetNameAllocator;
use Mnb\PHPExcel\Support\XlsxCompatibilityVerifier;
use Mnb\PHPExcel\Support\AdvancedWorkbookCapabilities;
use Mnb\PHPExcel\Support\BenchmarkSuite;
use Mnb\PHPExcel\Support\CompatibilityFixtureSuite;
use Mnb\PHPExcel\Support\DatabaseIntegrationSuite;
use Mnb\PHPExcel\Storage\StorageManager;
use Mnb\PHPExcel\Support\XlsxIntegrityValidator;
use PDO;
use Throwable;
use Mnb\PHPExcel\Validation\ArrayValidator;
use Mnb\PHPExcel\Validation\CustomValidatorRegistry;

final class MnbExcel
{
    public const VERSION = '1.6.0';

    private static ?ReaderRegistry $readerRegistry = null;
    private static ?DomainImportRegistry $domainImportRegistry = null;

    public static function version(): string
    {
        return self::VERSION;
    }

    public static function text(mixed $value): CellValue
    {
        return CellValue::text($value);
    }

    public static function number(int|float|string $value): CellValue
    {
        return CellValue::number($value);
    }

    public static function bool(bool $value): CellValue
    {
        return CellValue::bool($value);
    }

    /** @param array{format?:string} $options */
    public static function date(\DateTimeInterface|string $value, array $options = []): CellValue
    {
        return CellValue::date($value, $options);
    }

    /** @param array<string,mixed> $options */
    public static function formula(string $formula, mixed $cachedValue = null, array $options = []): CellValue
    {
        return CellValue::formula($formula, $cachedValue, $options);
    }

    public static function blank(): CellValue
    {
        return CellValue::blank();
    }

    /**
     * Scan source rows for formula-like text, unsafe formulas, invalid cell text,
     * long cell text, and long numeric strings that can lose precision in Excel.
     *
     * @param list<array<int|string,mixed>> $rows
     * @param array<string,mixed> $options
     * @return array{status:string,total_issues:int,issues:list<array<string,mixed>>}
     */
    public static function scanCells(array $rows, array $options = []): array
    {
        return (new CellSafetyScanner())->scan($rows, $options);
    }

    /**
     * Create a workbook builder from a single sheet array.
     *
     * @param array<int|string, mixed> $rows
     */
    public static function fromArray(array $rows): WorkbookBuilder
    {
        return WorkbookBuilder::fromArray($rows);
    }



    /**
     * Create a workbook builder from a JSON file.
     *
     * Supported JSON shapes: list of objects, single object, {"rows": [...]},
     * sheet map, {"sheets": {...}}, or {"sheets": [{"name": "...", "rows": [...]}]}.
     *
     * @param array<string,mixed> $options
     */
    public static function fromJson(string $path, array $options = []): WorkbookBuilder
    {
        $reader = new JsonReader();
        $workbook = $reader->readWorkbook($path, $options);
        $builder = WorkbookBuilder::fromWorkbookArray($workbook);

        if ((bool) ($options['with_header'] ?? true)) {
            $builder->withHeader();
        }

        return $builder;
    }

    /**
     * Create a report-ready workbook builder from array rows.
     *
     * @param array<int|string, mixed> $rows
     */
    public static function report(array $rows, string $template = 'business'): WorkbookBuilder
    {
        return WorkbookBuilder::fromArray($rows)
            ->withHeader()
            ->reportTemplate($template)
            ->freezeHeader()
            ->autoFilter();
    }

    /**
     * Create a workbook builder from many sheets.
     *
     * @param array<string, array<int|string, mixed>> $sheets
     */
    public static function fromWorkbookArray(array $sheets): WorkbookBuilder
    {
        return WorkbookBuilder::fromWorkbookArray($sheets);
    }


    /**
     * Create a standalone workbook containing import summary metrics.
     *
     * @param array<string,mixed> $summary
     * @param array<string,mixed> $validationResult
     */
    public static function fromImportSummary(array $summary, array $validationResult = []): WorkbookBuilder
    {
        return WorkbookBuilder::fromImportSummary($summary, $validationResult);
    }

    public static function safeFileName(string $filename, string $extension = ''): string
    {
        return WorkbookBuilder::safeFileName($filename, $extension);
    }

    public static function safeSheetName(string $name): string
    {
        return SheetNameAllocator::sanitize($name);
    }

    /**
     * Return resolved CSV dialect options, useful for UI/config previews.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function csvDialectOptions(array $options = []): array
    {
        return \Mnb\PHPExcel\Support\CsvDialect::resolve($options);
    }


    /**
     * Detect character encoding for CSV/text files without loading the full file.
     *
     * @param list<string>|null $candidates
     * @return array{encoding:string,confidence:float,source:string,bom:bool,notes:list<string>,sample_bytes:int,path:string}
     */
    public static function detectEncoding(string $path, int $sampleBytes = 65536, ?array $candidates = null): array
    {
        return EncodingDetector::detectFile($path, $sampleBytes, $candidates);
    }

    /**
     * Run local package-readiness checks before Git/Packagist release.
     *
     * @return array{status:string,checks:list<array{name:string,status:string,message:string}>,summary:array{passed:int,warning:int,failed:int}}
     */
    public static function releaseReadiness(string $root): array
    {
        return ReleaseReadiness::check($root);
    }

    /**
     * Report PHP extension/runtime capability for XLSX, SQL, encoding, and atomic-save workflows.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function environmentCheck(array $options = []): array
    {
        return EnvironmentDiagnostics::check($options);
    }

    /**
     * Return a public/developer alert when required XLSX/large-import extensions are missing.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function environmentAlert(array $options = []): array
    {
        return EnvironmentDiagnostics::alert($options);
    }

    /**
     * Return a ready-to-print environment alert message for CLI/admin screens.
     *
     * @param array<string,mixed> $options
     */
    public static function environmentAlertMessage(array $options = []): string
    {
        return EnvironmentDiagnostics::alertMessage($options);
    }

    /**
     * Run generated and optional external XLSX compatibility fixture verification.
     *
     * @param list<string> $fixturePaths
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function verifyXlsxCompatibility(array $fixturePaths = [], array $options = []): array
    {
        return (new XlsxCompatibilityVerifier())->verify($fixturePaths, $options);
    }

    /**
     * Return a reproducible benchmark plan for MNB PHPExcel and optional comparison libraries.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function benchmarkPlan(array $options = []): array
    {
        return BenchmarkSuite::plan($options);
    }

    /**
     * Run an internal low-memory large XLSX writer benchmark for this package.
     * Third-party comparison benchmarks should be run in a separate workspace with
     * those optional packages installed.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function benchmarkLargeWriter(int $rows, int $columns = 10, array $options = []): array
    {
        return BenchmarkSuite::runLargeWriter($rows, $columns, $options);
    }

    /** Return required real-world XLSX fixture coverage for release compatibility proof.
     *
     * @return array<string,mixed>
     */
    public static function compatibilityFixtureRequirements(): array
    {
        return CompatibilityFixtureSuite::requirements();
    }

    /** Verify real Excel/LibreOffice/Google Sheets/WPS fixture files in a directory.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function verifyCompatibilityFixtures(string $fixtureDir, array $options = []): array
    {
        return CompatibilityFixtureSuite::verify($fixtureDir, $options);
    }

    /** Return MySQL/PostgreSQL integration-test setup guidance.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function databaseIntegrationPlan(array $options = []): array
    {
        return DatabaseIntegrationSuite::plan($options);
    }

    /** Check whether optional MySQL/PostgreSQL integration-test connections are configured.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function databaseIntegrationCheck(array $options = []): array
    {
        return DatabaseIntegrationSuite::check($options);
    }

    /** Return current/planned advanced workbook capability matrix.
     *
     * @return array<string,mixed>
     */
    public static function advancedWorkbookCapabilities(): array
    {
        return AdvancedWorkbookCapabilities::matrix();
    }

    /**
     * Convert any caught exception into a public-safe response array.
     *
     * @return array<string,mixed>
     */
    public static function safeError(Throwable $throwable): array
    {
        return ErrorReporter::safe($throwable);
    }

    /**
     * Convert any caught exception into an error report. Pass debug=true only for logs/admin screens.
     *
     * @return array<string,mixed>
     */
    public static function errorReport(Throwable $throwable, bool $debug = false): array
    {
        return ErrorReporter::report($throwable, $debug);
    }

    /**
     * Read XLSX, CSV, JSON/JSON Lines, or XML using format auto-detection.
     * Pass ['format' => 'xlsx|csv|json|xml'] to override detection.
     *
     * @param array<string,mixed> $options
     */
    public static function read(string $path, array|ReaderOptions $options = []): ReadSession
    {
        $values = $options instanceof ReaderOptions ? $options->toArray() : ReaderOptions::fromArray($options)->toArray();
        return new ReadSession($path, self::readerRegistry()->resolve($path, $values), $values);
    }

    /** Instance-based entry point recommended for long-running workers. */
    public static function manager(): SpreadsheetManager
    {
        return SpreadsheetManager::create(ReaderRegistry::withBuiltIns());
    }

    public static function registerReaderPlugin(ReaderPluginInterface $plugin, int $priority = 0): void
    {
        self::readerRegistry()->registerPlugin($plugin, $priority);
    }

    public static function resetReaderRegistry(): void
    {
        self::$readerRegistry = null;
    }

    /** @param array<string,mixed> $options */
    public static function readXlsx(string $path, array $options = []): ReadSession
    {
        return new ReadSession($path, new XlsxReader(), $options);
    }

    /** @param array<string,mixed> $options */
    public static function readXml(string $path, array $options = []): ReadSession
    {
        return new ReadSession($path, new XmlReader(), $options);
    }

    /** @param array<string,mixed>|ReaderOptions $options */
    public static function readOds(string $path, array|ReaderOptions $options = []): ReadSession
    {
        return new ReadSession($path, new OdsReader(), $options);
    }

    /** Optional legacy XLS adapter; requires phpoffice/phpspreadsheet. */
    public static function readXls(string $path, array|ReaderOptions $options = []): ReadSession
    {
        return new ReadSession($path, new XlsReader(), $options);
    }

    /** @param array<string,mixed> $options */
    public static function detectFormat(string $path, array $options = []): string
    {
        return FileFormatDetector::detect($path, $options);
    }

    /**
     * Read a CSV file. Options become default read options for this session.
     *
     * @param array<string,mixed> $options
     */
    public static function readCsv(string $path, array $options = []): ReadSession
    {
        return new ReadSession($path, new CsvReader(), $options);
    }



    /**
     * Read a JSON file. Options become default read options for this session.
     *
     * @param array<string,mixed> $options
     */
    public static function readJson(string $path, array $options = []): ReadSession
    {
        return new ReadSession($path, new JsonReader(), $options);
    }




    /**
     * Resolve database settings from .env, PHP config file, constants, runtime env, DSN, or array.
     *
     * @param array<string,mixed>|string|null $source
     * @param array<string,mixed> $overrides
     * @return array<string,mixed>
     */
    public static function dbConfig(array|string|null $source = null, array $overrides = []): array
    {
        return DatabaseConfigResolver::resolve($source, $overrides);
    }

    /**
     * Create a PDO connection from .env, PHP config file, constants, runtime env, DSN, array, or existing PDO.
     *
     * @param PDO|array<string,mixed>|string|null $source
     * @param array<string,mixed> $overrides
     */
    public static function pdo(PDO|array|string|null $source = null, array $overrides = []): PDO
    {
        return DatabaseConnectionFactory::make($source, $overrides);
    }

    /** Alias of pdo() for readability in application code.
     *
     * @param PDO|array<string,mixed>|string|null $source
     * @param array<string,mixed> $overrides
     */
    public static function databaseConnection(PDO|array|string|null $source = null, array $overrides = []): PDO
    {
        return self::pdo($source, $overrides);
    }

    /**
     * Show which .env/config paths would be checked for database settings.
     *
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function databaseConfigSummary(array $options = []): array
    {
        return DatabaseConfigResolver::summary($options);
    }

    /**
     * Analyze an XLSX file without loading all rows. Use this before importing large workbooks.
     *
     * @param array<string,mixed> $options accurate_row_count, scan_features, time_budget_seconds, sheet, server
     * @return array<string,mixed>
     */
    public static function analyzeXlsxForImport(string $path, array $options = []): array
    {
        return (new LargeExcelPreflightAnalyzer())->analyze($path, $options);
    }

    /**
     * Recommend the safest import method for an XLSX file and server profile.
     *
     * @param array<string,mixed> $serverOptions server, memory_limit, max_execution_time, allow_http_large_import
     * @param array<string,mixed> $analysisOptions accurate_row_count, scan_features, time_budget_seconds, sheet
     * @return array<string,mixed>
     */
    public static function recommendImportMethod(string $path, array $serverOptions = [], array $analysisOptions = []): array
    {
        $profile = (new LargeExcelPreflightAnalyzer())->analyze($path, array_merge($analysisOptions, ['server' => $serverOptions]));
        return (new ImportMethodAdvisor())->recommendFromProfile($profile, $serverOptions);
    }

    /**
     * Recommend an import method from already-known metrics. Useful for tests, UIs, and custom analyzers.
     *
     * @param array<string,mixed> $profile
     * @param array<string,mixed> $serverOptions
     * @return array<string,mixed>
     */
    public static function recommendImportMethodFromProfile(array $profile, array $serverOptions = []): array
    {
        return (new ImportMethodAdvisor())->recommendFromProfile($profile, $serverOptions);
    }

    /**
     * Return the method decision matrix used by the advisor.
     *
     * @return array<string,mixed>
     */
    public static function importMethodMatrix(): array
    {
        return (new ImportMethodAdvisor())->decisionMatrix();
    }

    /**
     * Create a streaming XLSX read session for large imports. It never loads the whole workbook.
     *
     * @param array<string,mixed> $options sheet, with_header, skip_rows, only_columns, time_budget_seconds
     */
    public static function largeRead(string $path, array $options = []): LargeXlsxReadSession
    {
        return new LargeXlsxReadSession($path, null, $options);
    }


    /**
     * Create a streaming XLSX/CSV-ZIP export session for very large exports.
     * The iterable is consumed once and rows are never buffered as a full workbook.
     *
     * @param iterable<array<int|string,mixed>> $rows
     * @param array<string,mixed> $options
     */
    public static function largeExport(iterable $rows, array $options = []): LargeXlsxWriteSession
    {
        return new LargeXlsxWriteSession($rows, null, $options);
    }

    /**
     * Stream rows directly to a large XLSX file.
     *
     * @param iterable<array<int|string,mixed>> $rows
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function largeWrite(string $path, iterable $rows, array $options = []): array
    {
        return self::largeExport($rows, $options)->save($path);
    }

    /**
     * Stream rows directly to a split CSV ZIP fallback. Useful when XLSX is too large
     * for practical Excel opening or row limits.
     *
     * @param iterable<array<int|string,mixed>> $rows
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function largeWriteCsvZip(string $path, iterable $rows, array $options = []): array
    {
        return self::largeExport($rows, $options)->saveCsvZip($path);
    }

    /**
     * Create a streaming export session from a PDO cursor. The query result is not
     * loaded into memory.
     *
     * @param array<int|string,mixed> $params
     * @param array<string,mixed> $options
     */
    public static function largeExportFromSql(PDO|array|string|null $pdo, string $query, array $params = [], array $options = []): LargeXlsxWriteSession
    {
        $connection = DatabaseConnectionFactory::make($pdo, is_array($options['db'] ?? null) ? $options['db'] : []);
        return self::largeExport(LargePdoCursor::rows($connection, $query, $params, $options), $options)->withHeader((bool) ($options['with_header'] ?? true));
    }

    /**
     * Stream a PDO cursor directly to a large XLSX file.
     *
     * @param array<int|string,mixed> $params
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function largeWriteFromSql(string $path, PDO|array|string|null $pdo, string $query, array $params = [], array $options = []): array
    {
        return self::largeExportFromSql($pdo, $query, $params, $options)->save($path);
    }

    /**
     * Build an auto-import plan: preflight profile + selected method + ready-to-use routing advice.
     *
     * @param array<string,mixed> $serverOptions
     * @param array<string,mixed> $analysisOptions
     * @return array<string,mixed>
     */
    public static function autoImportPlan(string $path, array $serverOptions = [], array $analysisOptions = []): array
    {
        $profile = (new LargeExcelPreflightAnalyzer())->analyze($path, array_merge($analysisOptions, ['server' => $serverOptions]));
        $advice = (new ImportMethodAdvisor())->recommendFromProfile($profile, $serverOptions);

        return [
            'status' => 'ok',
            'profile' => $profile,
            'advice' => $advice,
            'selected_method' => $advice['method'],
            'chunk_size' => $advice['recommended_chunk_size'],
            'route' => $advice['recommended_route'],
        ];
    }


    /**
     * Stream a large XLSX file into a SQL table using chunk validation, PDO batch inserts,
     * failed-row CSV export, and a resumable JSON progress manifest.
     *
     * @param array<string,mixed> $options sheet, with_header, rules, batch_size, chunk_size, manifest_path, failed_rows_csv, resume, time_budget_seconds
     * @return array<string,mixed>
     */
    public static function largeImportToSql(string $path, PDO|array|string|null $pdo, string $table, array $options = []): array
    {
        return (new LargeExcelDatabaseImportEngine())->importToSql($path, DatabaseConnectionFactory::make($pdo, is_array($options['db'] ?? null) ? $options['db'] : []), $table, $options);
    }

    /**
     * Import all or selected workbook sheets one by one using the large database import engine.
     *
     * @param string|array<int|string,string> $tableMap String table prefix/name or sheet-name/index => table map.
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function largeImportWorkbookToSql(string $path, PDO|array|string|null $pdo, string|array $tableMap, array $options = []): array
    {
        return (new LargeExcelDatabaseImportEngine())->importWorkbookToSql($path, DatabaseConnectionFactory::make($pdo, is_array($options['db'] ?? null) ? $options['db'] : []), $tableMap, $options);
    }

    /**
     * Return the default resumable manifest path for a large XLSX import job.
     */
    public static function largeImportManifestPath(string $path, string $table = 'import', int|string $sheet = 1): string
    {
        return LargeImportManifest::defaultPath($path, $table, $sheet);
    }



    /**
     * Configure application storage paths for uploads, manifests, failed rows, reports, and temp files.
     *
     * @param array<string,string> $paths
     * @return array<string,string>
     */
    public static function storage(array $paths = []): array
    {
        return StorageManager::configure($paths);
    }

    /** Return a configured storage path by key. */
    public static function storagePath(string $key, string $filename = ''): string
    {
        return StorageManager::path($key, $filename);
    }

    /**
     * Load import profiles from an array or PHP config file.
     *
     * @param array<string,array<string,mixed>>|string $profiles
     */
    public static function loadImportProfiles(array|string $profiles): void
    {
        ImportProfileManager::load($profiles);
    }

    /** Register one reusable import profile. */
    public static function registerImportProfile(string $name, array $profile): void
    {
        ImportProfileManager::register($name, $profile);
    }

    /** Start a reusable import profile. */
    public static function profile(string $name): ImportProfile
    {
        return ImportProfileManager::profile($name);
    }

    /** @return array<string,array<string,mixed>> */
    public static function importProfiles(): array
    {
        return ImportProfileManager::all();
    }

    /** Read a progress/status manifest for admin dashboards and job runners. */
    public static function importStatus(string $manifestPath, array $options = []): array
    {
        return ImportStatusReader::read($manifestPath, $options);
    }


    /** Build an admin/API friendly dashboard response from an import manifest. */
    public static function importDashboard(string|array $manifestOrStatus, array $options = []): array
    {
        return ImportDashboardHelper::response($manifestOrStatus, $options);
    }

    /** Alias for importDashboard(), useful in controllers returning JSON. */
    public static function importStatusResponse(string|array $manifestOrStatus, array $options = []): array
    {
        return self::importDashboard($manifestOrStatus, $options);
    }

    /** Resume an existing large import job from its manifest. */
    public static function resumeImport(string $manifestPath, PDO|array|string|null $pdo = null, array $options = []): array
    {
        return ImportJobRunner::resume($manifestPath, $pdo, $options);
    }

    /** Validate an uploaded file path or $_FILES-style array before import. */
    public static function validateUpload(array|string $file, array $options = []): array
    {
        return UploadSafetyValidator::validate($file, $options);
    }

    /** Store a validated $_FILES-style upload and return an AJAX/API-ready response. */
    public static function handleAjaxUpload(array|string $file, array $options = []): array
    {
        return (new AjaxUploadHandler())->handle($file, $options);
    }

    /** Framework-neutral spreadsheet API dispatcher for upload, preview, import, status, and export actions. */
    public static function api(string $action, array $request, PDO|array|string|null $pdo = null): array
    {
        return (new SpreadsheetApi())->handle($action, $request, $pdo);
    }

    /** Complete HTTP endpoint with routing, authentication, CORS and rate limiting. @param array<string,mixed> $options */
    public static function apiHttp(array $options = [], PDO|array|string|null $pdo = null): HttpResponse
    {
        return (new SpreadsheetHttpEndpoint(new SpreadsheetApi(), $options))->handleGlobals($pdo);
    }

    /** Generate a dependency-free browser upload form and client. @param array<string,mixed> $options */
    public static function ajaxUploader(string $endpoint, array $options = []): string
    {
        return AjaxUploader::html($endpoint, $options);
    }

    /** Import multiple spreadsheet files into one or dynamically resolved SQL table. */
    public static function importFilesToSql(array $files, PDO|array|string|null $pdo, string $table, array $options = []): array
    {
        return (new MultiFileImportManager())->importToSql($files, $pdo, $table, $options);
    }

    /** Import multiple files through one typed domain preset. */
    public static function importDomainFiles(DomainImportType|string $domain, array $files, PDO|array|string|null $pdo = null, string $table = '', array $options = []): array
    {
        return (new MultiFileImportManager())->importDomain($domain, $files, $pdo, $table, $options);
    }

    /** Create the built-in durable filesystem queue. */
    public static function queue(string $directory): SpreadsheetQueue
    {
        return new SpreadsheetQueue(new FileQueue($directory));
    }

    /** Create a spreadsheet queue from any backend implementation. */
    public static function queueBackend(QueueBackendInterface $backend): SpreadsheetQueue
    {
        return new SpreadsheetQueue($backend);
    }

    /** Create a transactional multi-host PDO queue. */
    public static function pdoQueue(PDO $pdo, string $table = 'mnb_excel_queue'): SpreadsheetQueue
    {
        return new SpreadsheetQueue(new PdoQueue($pdo, $table));
    }

    /** Process queued spreadsheet jobs without requiring a framework queue package. */
    public static function workQueue(string $directory, array $options = []): array
    {
        return self::queue($directory)->work($options);
    }

    /** Send a generated or existing spreadsheet as a MIME email attachment. */
    public static function emailGeneratedExcel(WorkbookBuilder|string $workbook, string|array $to, string $subject, string $body = '', array $options = []): bool
    {
        $mailer = $options['mailer'] ?? null;
        if ($mailer instanceof MailerInterface) {
            $resolved = $mailer;
        } elseif (is_callable($mailer)) {
            $resolved = new CallbackMailer($mailer);
        } elseif (is_array($options['smtp'] ?? null)) {
            $resolved = new SmtpMailer((array) $options['smtp']);
        } else {
            $resolved = new NativeMailer(
                is_callable($options['transport'] ?? null) ? $options['transport'] : null,
                (string) ($options['from'] ?? '')
            );
        }
        unset($options['mailer'], $options['transport'], $options['smtp'], $options['from']);
        return (new SpreadsheetMailer($resolved))->send($workbook, $to, $subject, $body, $options);
    }

    /** Create a persistent cron-style import/export scheduler. */
    public static function scheduler(string $storePath): SpreadsheetScheduler
    {
        return new SpreadsheetScheduler(new FileScheduler($storePath));
    }

    /** Create a transactional multi-host PDO scheduler. */
    public static function pdoScheduler(PDO $pdo, string $table = 'mnb_excel_schedule'): SpreadsheetScheduler
    {
        return new SpreadsheetScheduler(new PdoScheduler($pdo, $table));
    }

    /** Run a scheduler continuously with process locking and graceful shutdown. @param array<string,mixed> $options */
    public static function runScheduler(SpreadsheetScheduler $scheduler, array $options = []): array
    {
        return (new SchedulerRunner($scheduler))->runForever($options);
    }

    /** @param callable(array<string,mixed>):mixed $listener */
    public static function on(string $event, callable $listener): void
    {
        EventDispatcher::listen($event, $listener);
    }

    /** @param array<string,mixed> $payload @return array<string,mixed> */
    public static function dispatch(string $event, array $payload = []): array
    {
        return EventDispatcher::dispatch($event, $payload);
    }

    /** Register a plugin object or closure. */
    public static function plugin(MnbExcelPluginInterface|callable $plugin): void
    {
        PluginManager::register($plugin);
    }

    /** @return list<string> */
    public static function plugins(): array
    {
        return PluginManager::plugins();
    }

    /** @param callable(array<string|int,mixed>, array<string,mixed>):array<string|int,mixed> $transformer */
    public static function transformer(string $name, callable $transformer): void
    {
        RowTransformerPipeline::register($name, $transformer);
    }

    /** @param callable(mixed,array<string,mixed>,array<string,mixed>):bool|string|null $callback */
    public static function validator(string $name, callable $callback): void
    {
        CustomValidatorRegistry::register($name, $callback);
    }

    /** Optional PSR-3-like logger or callable logger. */
    public static function setLogger(mixed $logger): void
    {
        LoggerBridge::set($logger);
    }

    /**
     * Validate an XLSX package for required parts, relationships, content types,
     * XML well-formedness, and worksheet relationship ID consistency.
     *
     * @param array<string,mixed> $options
     * @return array{status:string,valid:bool,errors:list<string>,warnings:list<string>,checks:list<array{name:string,status:string,message:string}>,summary:array{passed:int,warning:int,failed:int},path:string}
     */
    public static function validateXlsx(string $path, array $options = []): array
    {
        return (new XlsxIntegrityValidator())->validate($path, $options);
    }

    /**
     * Throw MnbExcelException when XLSX integrity validation fails.
     *
     * @param array<string,mixed> $options
     */
    public static function assertValidXlsx(string $path, array $options = []): void
    {
        (new XlsxIntegrityValidator())->assertValid($path, $options);
    }

    /**
     * Inspect an XLSX workbook structure without converting workbook rows to arrays.
     *
     * @return array<string,mixed>
     */
    public static function inspect(string $path): array
    {
        return (new XlsxInspector())->inspect($path);
    }

    /** @return list<string> */
    public static function sheetNames(string $path): array
    {
        return (new XlsxInspector())->sheetNames($path);
    }

    /**
     * Export SQL query result to an array workbook builder.
     *
     * @param array<int|string, mixed> $params
     */
    public static function fromSql(PDO|array|string|null $pdo, string $query, array $params = [], array $dbOptions = []): WorkbookBuilder
    {
        $pdo = DatabaseConnectionFactory::make($pdo, $dbOptions);
        try {
            $stmt = $pdo->prepare($query);
            if ($stmt === false) {
                throw MnbExcelException::withCode('Unable to prepare SQL export query.', ErrorCode::SQL_EXPORT_FAILED);
            }

            $stmt->execute($params);

            /** @var array<int, array<string, mixed>> $rows */
            $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

            return WorkbookBuilder::fromArray($rows);
        } catch (MnbExcelException $e) {
            throw $e;
        } catch (Throwable $e) {
            throw MnbExcelException::withCode(
                'SQL export failed: ' . $e->getMessage(),
                ErrorCode::SQL_EXPORT_FAILED,
                ['query' => $query],
                $e
            );
        }
    }


    /**
     * Preview import quality before validation or SQL insert.
     *
     * @param list<array<string,mixed>|list<mixed>> $rows
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public static function previewImport(array $rows, array $options = []): array
    {
        return (new ImportQualityAnalyzer())->preview($rows, $options);
    }

    /**
     * Suggest source-to-target column mappings for import screens.
     *
     * @param list<string> $sourceColumns
     * @param list<string> $targetColumns
     * @param array<string,list<string>|string> $aliases
     * @return array<string,array{target:?string,confidence:float,reason:string}>
     */
    public static function suggestColumnMap(array $sourceColumns, array $targetColumns, array $aliases = [], float $minConfidence = 0.55): array
    {
        return (new ImportQualityAnalyzer())->suggestColumnMap($sourceColumns, $targetColumns, $aliases, $minConfidence);
    }

    /**
     * Detect duplicate rows by one or more columns.
     *
     * @param list<array<string,mixed>|list<mixed>> $rows
     * @param list<string> $columns
     * @param array<string,mixed> $options
     * @return list<array{key:string,count:int,rows:list<int>}>
     */
    public static function duplicateRows(array $rows, array $columns, array $options = []): array
    {
        return (new ImportQualityAnalyzer())->findDuplicates($rows, $columns, $options);
    }

    /**
     * Validate associative array rows using MNB PHPExcel validation rules.
     *
     * @param list<array<string, mixed>> $rows
     * @param array<string, string> $rules
     * @param array<string,mixed> $options
     * @return array{valid:list<array<string,mixed>>,failed:list<array{row:int,errors:list<string>,data:array<string,mixed>}>}
     */
    public static function validateArray(array $rows, array $rules, array $options = []): array
    {
        return (new ArrayValidator())->validate($rows, $rules, $options);
    }

    /**
     * Import-focused alias for validateArray().
     *
     * @param list<array<string, mixed>> $rows
     * @param array<string, string> $rules
     * @param array<string,mixed> $options
     * @return array{valid:list<array<string,mixed>>,failed:list<array{row:int,errors:list<string>,data:array<string,mixed>}>}
     */
    public static function validateImport(array $rows, array $rules, array $options = []): array
    {
        return self::validateArray($rows, $rules, $options);
    }

    /**
     * Build an Excel/CSV report from failed validation rows.
     *
     * @param list<array{row:int,errors:list<string>,data:array<string,mixed>}> $failedRows
     */
    public static function fromFailedRows(array $failedRows): WorkbookBuilder
    {
        $rows = [];
        foreach ($failedRows as $failed) {
            $rows[] = array_merge([
                'row_number' => $failed['row'] ?? null,
                'errors' => implode('; ', $failed['errors'] ?? []),
            ], is_array($failed['data'] ?? null) ? $failed['data'] : []);
        }

        return WorkbookBuilder::fromArray($rows)->withHeader()->styleHeader([
            'font' => ['bold' => true, 'color' => '#FFFFFF'],
            'fill' => '#B42318',
            'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap_text' => true],
            'border' => true,
        ]);
    }

    /** @return array<string,array<string,mixed>> */
    public static function domainImportSchemas(): array
    {
        return self::domainImporter()->schemas();
    }

    /** @return array<string,mixed> */
    public static function domainImportSchema(DomainImportType|string $domain): array
    {
        return self::domainImporter()->schema($domain);
    }

    /** Replace one built-in preset for this process. */
    public static function registerDomainImportPreset(DomainImportPreset $preset): void
    {
        self::domainImportRegistry()->register($preset);
    }

    /** @param array<string,mixed> $options */
    public static function domainImportTemplate(DomainImportType|string $domain, array $options = []): WorkbookBuilder
    {
        $preset = self::domainImportRegistry()->get($domain);
        return WorkbookBuilder::importTemplate($preset->templateColumns(), array_replace([
            'title' => ucwords(str_replace('_', ' ', $preset->type->value)) . ' Import Template',
            'instructions' => 'Complete the required columns. Header aliases are accepted during import.',
            'sample_rows' => 1,
        ], $options));
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public static function previewDomainImport(DomainImportType|string $domain, string $path, array $options = []): array
    {
        return self::domainImporter()->preview($domain, $path, $options);
    }

    /** @param PDO|array<string,mixed>|string|null $pdo @param array<string,mixed> $options @return array<string,mixed> */
    public static function importDomain(DomainImportType|string $domain, string $path, PDO|array|string|null $pdo = null, string $table = '', array $options = []): array
    {
        return self::domainImporter()->import($domain, $path, $pdo, $table, $options);
    }

    public static function importUsers(string $path, PDO|array|string|null $pdo = null, string $table = 'users', array $options = []): array { return self::importDomain(DomainImportType::Users, $path, $pdo, $table, $options); }
    public static function importProducts(string $path, PDO|array|string|null $pdo = null, string $table = 'products', array $options = []): array { return self::importDomain(DomainImportType::Products, $path, $pdo, $table, $options); }
    public static function importOrders(string $path, PDO|array|string|null $pdo = null, string $table = 'orders', array $options = []): array { return self::importDomain(DomainImportType::Orders, $path, $pdo, $table, $options); }
    public static function importInventory(string $path, PDO|array|string|null $pdo = null, string $table = 'inventory', array $options = []): array { return self::importDomain(DomainImportType::Inventory, $path, $pdo, $table, $options); }
    public static function importStudents(string $path, PDO|array|string|null $pdo = null, string $table = 'students', array $options = []): array { return self::importDomain(DomainImportType::Students, $path, $pdo, $table, $options); }
    public static function importAttendance(string $path, PDO|array|string|null $pdo = null, string $table = 'attendance', array $options = []): array { return self::importDomain(DomainImportType::Attendance, $path, $pdo, $table, $options); }
    public static function importMarks(string $path, PDO|array|string|null $pdo = null, string $table = 'marks', array $options = []): array { return self::importDomain(DomainImportType::Marks, $path, $pdo, $table, $options); }
    public static function importContacts(string $path, PDO|array|string|null $pdo = null, string $table = 'contacts', array $options = []): array { return self::importDomain(DomainImportType::Contacts, $path, $pdo, $table, $options); }
    public static function importLocations(string $path, PDO|array|string|null $pdo = null, string $table = 'locations', array $options = []): array { return self::importDomain(DomainImportType::Locations, $path, $pdo, $table, $options); }
    public static function importBlogPosts(string $path, PDO|array|string|null $pdo = null, string $table = 'blog_posts', array $options = []): array { return self::importDomain(DomainImportType::BlogPosts, $path, $pdo, $table, $options); }
    public static function importImagesWithPaths(string $path, PDO|array|string|null $pdo = null, string $table = 'media', array $options = []): array { return self::importDomain(DomainImportType::Media, $path, $pdo, $table, $options); }
    public static function importMedia(string $path, PDO|array|string|null $pdo = null, string $table = 'media', array $options = []): array { return self::importImagesWithPaths($path, $pdo, $table, $options); }
    public static function importCategories(string $path, PDO|array|string|null $pdo = null, string $table = 'categories', array $options = []): array { return self::importDomain(DomainImportType::Categories, $path, $pdo, $table, $options); }

    private static function domainImporter(): DomainImporter
    {
        return DomainImporter::create(self::domainImportRegistry(), self::readerRegistry());
    }

    private static function domainImportRegistry(): DomainImportRegistry
    {
        return self::$domainImportRegistry ??= DomainImportRegistry::withBuiltIns();
    }

    private static function readerRegistry(): ReaderRegistry
    {
        return self::$readerRegistry ??= ReaderRegistry::withBuiltIns();
    }

}
