# mnb/mnb-phpexcel-application

Legacy facade, workbook builder, application helpers, plugins, diagnostics, and full workflow API for MNB PHPExcel.

This package is generated from the MNB PHPExcel monorepo. Do not copy source files between modules manually.

## Install

```bash
composer require mnb/mnb-phpexcel-application
```

See the main project documentation for typed options, streaming reads, and compatibility notes.

## XLSX lightweight information facade

The application facade delegates lightweight XLSX inspection to
`mnb/mnb-phpexcel-xlsx`:

```php
use Mnb\PHPExcel\MnbExcel;

$file = MnbExcel::fileInfo('orders.xlsx');
$sheets = MnbExcel::sheetsInfo('orders.xlsx');
$rows = MnbExcel::rowCount('orders.xlsx', 'Orders');
```

The implementation remains in the XLSX package; the core package is unchanged.

