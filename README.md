# MNB PHPExcel Application

Complete facade, workbook builder, diagnostics, plugins, queues, HTTP helpers, mail workflows, and cross-format application API.

```bash
composer require mnb/mnb-phpexcel-application:^2.0
```

```php
use Mnb\PHPExcel\MnbExcel;

$rows = MnbExcel::read('report.xlsx')
    ->sheet('Data')
    ->withHeaderRow()
    ->toArray();

MnbExcel::fromArray($rows)
    ->withHeader()
    ->freezeHeader()
    ->autoWidth()
    ->save('report-copy.xlsx');
```

The application package installs every format and integration module at the same `^2.0` release line, including native XLS and the optional XLSX/database bridge.
