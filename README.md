# MNB PHPExcel Application

Complete facade, workbook builder, diagnostics, plugins, queues, HTTP helpers, mail workflows, and cross-format application API.
## MNB PHPExcel Assistant

Generate MNB PHPExcel code using our dedicated ChatGPT assistant:

[Open MNB PHPExcel AI Assistant](https://chatgpt.com/g/g-6a6e31d80350819194b68853d41c1561-mnb-phpexcel-assistant)
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
