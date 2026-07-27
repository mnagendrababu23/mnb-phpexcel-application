<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorkbookFactory;
use Mnb\PHPExcel\MnbExcel;
use Mnb\PHPExcel\Writer\XlsxWriter;

$directory = sys_get_temp_dir() . '/mnb-app-quick-info-' . bin2hex(random_bytes(5));
if (!mkdir($directory, 0775, true) && !is_dir($directory)) {
    throw new RuntimeException('Unable to create test directory.');
}
$path = $directory . '/quick-info.xlsx';

try {
    $sheet = WorkbookFactory::worksheet([
        ['ID' => 1, 'Name' => 'Alpha'],
        ['ID' => 2, 'Name' => 'Beta'],
    ], 'Data', true);
    (new XlsxWriter())->write(new WorkbookData([$sheet]), $path);

    assert(MnbExcel::fileInfo($path)['sheet_count'] === 1);
    assert(MnbExcel::sheetInfo($path, 'Data')['declared_last_row'] === 3);
    assert(MnbExcel::rowCount($path, 'Data') === 3);
    assert(MnbExcel::rowCounts($path) === ['Data' => 3]);

    echo "xlsx_quick_info_facade_smoke passed\n";
} finally {
    @unlink($path);
    @rmdir($directory);
}
