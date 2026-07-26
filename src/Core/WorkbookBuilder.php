<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Core;

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Import\SqlImporter;
use Mnb\PHPExcel\Security\CellSafetyScanner;
use Mnb\PHPExcel\Security\FormulaGuard;
use Mnb\PHPExcel\Support\Arr;
use Mnb\PHPExcel\Support\DatabaseConnectionFactory;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\SheetNameAllocator;
use Mnb\PHPExcel\Support\ValueSanitizer;
use Mnb\PHPExcel\Writer\CsvWriter;
use Mnb\PHPExcel\Writer\JsonWriter;
use Mnb\PHPExcel\Writer\XmlWriter;
use Mnb\PHPExcel\Writer\XlsxWriter;
use PDO;

final class WorkbookBuilder
{
    /** @var array<string, array<int|string, mixed>> */
    private array $sourceSheets;

    /** @var array<int|string, string> */
    private array $columns = [];

    private bool $withHeader = false;

    /** @var list<string> */
    private array $textColumns = [];

    /** @var array<string, string> */
    private array $dateColumns = [];

    /** @var list<string> */
    private array $numberColumns = [];

    private bool $freezeHeader = false;
    private bool $autoFilter = false;
    private int $freezeRows = 0;
    private int $freezeColumns = 0;
    private ?string $freezeTopLeftCell = null;
    private ?string $autoFilterRange = null;

    /** @var list<array<string,mixed>> */
    private array $filterColumns = [];

    /** @var list<array<string,mixed>> */
    private array $nativeConditionalFormats = [];

    /** @var list<array<string,mixed>> */
    private array $dataValidations = [];

    /** @var array<string,list<array<string,mixed>>> */
    private array $charts = [];
    private bool $escapeFormulaLikeText = true;

    /** @var array<string, mixed>|string */
    private array|string $headerStyle = [];

    /** @var list<string> */
    private array $mergeCells = [];

    /** @var array<int|string, float|int> */
    private array $columnWidths = [];

    /** @var array<int, float|int> */
    private array $rowHeights = [];

    /** @var array<string,list<array{path:string,cell:string,width?:int,height?:int,name?:string}>> */
    private array $images = [];

    /** @var array<string, list<array{cell:string,url:string,display?:string,tooltip?:string}>> */
    private array $hyperlinks = [];

    /** @var array<string, list<array{cell:string,author:string,text:string,width?:float|int,height?:float|int,visible?:bool}>> */
    private array $comments = [];

    /** @var list<array{values:list<mixed>,style?:string|array<string,mixed>,height?:float|int,merge?:bool}> */
    private array $titleRows = [];

    /** @var list<array{values:list<mixed>,style?:string|array<string,mixed>,height?:float|int}> */
    private array $summaryRows = [];

    /** @var list<array{values:list<mixed>,style?:string|array<string,mixed>,height?:float|int,merge?:bool}> */
    private array $footerRows = [];

    /** @var array<string, array<string,mixed>> */
    private array $namedStyles = [];

    /** @var array<int|string, string|array<string,mixed>> */
    private array $columnStyles = [];

    /** @var array<int, string|array<string,mixed>> */
    private array $rowStyles = [];

    /** @var array<string, string|array<string,mixed>> */
    private array $cellStyles = [];

    /** @var array<string, string|array<string,mixed>> */
    private array $rangeStyles = [];

    private bool $autoWidth = false;

    /** @var array{min:int,max:int,padding:int} */
    private array $autoWidthOptions = ['min' => 8, 'max' => 45, 'padding' => 2];

    /** @var list<array{condition:callable,style:string|array<string,mixed>}> */
    private array $conditionalRowStyles = [];

    /** @var array<string,mixed> */
    private array $metadata = [];

    /** @var array<string,mixed> */
    private array $advancedObjectPreservation = [];

    /** @var array<string,mixed> */
    private array $xlsxIntegrityValidation = ['enabled' => true];

    /** @var array{name:string,rows:list<array<int,mixed>>,style?:string|array<string,mixed>}|null */
    private ?array $importSummarySheet = null;

    /** @var array<string,mixed> */
    private array $csvOptions = [
        'dialect' => 'excel',
        'bom' => true,
        'injection_policy' => 'escape',
    ];

    /** @var array<string,mixed> */
    private array $cellSafetyOptions = [
        'max_text_length' => 32767,
        'long_text_policy' => 'truncate',
        'control_char_policy' => 'strip',
        'large_number_as_text' => true,
        'allow_precision_loss' => false,
    ];

    /** @var array<string,mixed> */
    private array $formulaOptions = [
        'formula_policy' => 'safe',
    ];

    /**
     * @param array<string, array<int|string, mixed>> $sourceSheets
     */
    private function __construct(array $sourceSheets)
    {
        $this->sourceSheets = $sourceSheets;
        $this->registerDefaultReportStyles();
    }

    /** @param array<int|string, mixed> $rows */
    public static function fromArray(array $rows): self
    {
        return new self(['Sheet1' => $rows]);
    }

    /** @param array<string, array<int|string, mixed>> $sheets */
    public static function fromWorkbookArray(array $sheets): self
    {
        if ($sheets === []) {
            throw new MnbExcelException('Workbook array must contain at least one sheet.');
        }

        return new self($sheets);
    }

    /**
     * Build a one-sheet import summary workbook.
     *
     * @param array<string,mixed> $summary
     * @param array<string,mixed> $validationResult
     */
    public static function fromImportSummary(array $summary, array $validationResult = []): self
    {
        $builder = new self(['Import Summary' => []]);
        $builder->sourceSheets = ['Import Summary' => $builder->buildImportSummaryRows($summary, $validationResult)];
        return $builder
            ->rowStyle(1, 'mnb.header.blue')
            ->autoWidth(['min' => 12, 'max' => 55, 'padding' => 3])
            ->freezeHeader()
            ->autoFilter()
            ->metadata([
                'title' => 'Import Summary',
                'subject' => 'MNB PHPExcel import summary report',
            ]);
    }

    public function withHeader(bool $enabled = true): self
    {
        $this->withHeader = $enabled;
        return $this;
    }

    /** @param array<int|string, string> $columns */
    public function columns(array $columns): self
    {
        $this->columns = $columns;
        return $this;
    }

    /** @param list<string> $columns */
    public function textColumns(array $columns): self
    {
        $this->textColumns = array_values($columns);
        return $this;
    }

    /** @param array<string, string> $columns */
    public function dateColumns(array $columns): self
    {
        $this->dateColumns = $columns;
        return $this;
    }

    /** @param list<string> $columns */
    public function numberColumns(array $columns): self
    {
        $this->numberColumns = array_values($columns);
        return $this;
    }

    public function freezeHeader(bool $enabled = true): self
    {
        $this->freezeHeader = $enabled;
        return $this;
    }

    public function autoFilter(bool $enabled = true): self
    {
        $this->autoFilter = $enabled;
        return $this;
    }


    /** Freeze arbitrary rows and columns. Example: freezePanes(1, 2) freezes row 1 and columns A:B. */
    public function freezePanes(int $rows = 1, int $columns = 0, ?string $topLeftCell = null): self
    {
        if ($rows < 0 || $columns < 0) {
            throw new MnbExcelException('Freeze rows and columns cannot be negative.');
        }
        if ($topLeftCell !== null) {
            $topLeftCell = $this->normalizeCellReference($topLeftCell, 'freeze pane');
        }
        $this->freezeRows = $rows;
        $this->freezeColumns = $columns;
        $this->freezeTopLeftCell = $topLeftCell;
        $this->freezeHeader = false;
        return $this;
    }

    /** Freeze all rows above and columns left of the selected cell. */
    public function freezeAt(string $cell): self
    {
        $cell = $this->normalizeCellReference($cell, 'freeze pane');
        [$column, $row] = Coordinate::splitCellRef($cell);
        return $this->freezePanes(max(0, $row - 1), max(0, $column - 1), $cell);
    }

    /** Configure an explicit auto-filter range such as A1:H500. */
    public function autoFilterRange(string $range): self
    {
        $this->autoFilterRange = $this->normalizeRangeReference($range, 'auto-filter');
        $this->autoFilter = true;
        return $this;
    }

    /**
     * Add a native Excel filter definition.
     *
     * Supported types: values, custom, top10, dynamic, color.
     * @param array<string,mixed> $criteria
     */
    public function filterColumn(int|string $column, array $criteria): self
    {
        $index = is_int($column) || ctype_digit((string) $column)
            ? (int) $column
            : Coordinate::columnNameToIndex((string) $column);
        if ($index < 1) {
            throw new MnbExcelException('Filter column must be a positive index or Excel column name.');
        }
        $type = strtolower((string) ($criteria['type'] ?? 'values'));
        if (!in_array($type, ['values', 'custom', 'top10', 'dynamic', 'color'], true)) {
            throw new MnbExcelException('Unsupported filter type: ' . $type);
        }
        $criteria['type'] = $type;
        $criteria['column'] = $index;
        $this->filterColumns[] = $criteria;
        $this->autoFilter = true;
        return $this;
    }

    /** @param list<mixed> $values */
    public function filterValues(int|string $column, array $values, bool $includeBlank = false): self
    {
        return $this->filterColumn($column, ['type' => 'values', 'values' => array_values($values), 'include_blank' => $includeBlank]);
    }

    /** Add a native conditional-formatting rule to an XLSX range. */
    public function conditionalFormatting(string $range, string $type, array $options = []): self
    {
        $range = $this->normalizeRangeReference($range, 'conditional formatting');
        $type = strtolower(trim($type));
        if (!in_array($type, ['cell_is', 'expression', 'color_scale', 'data_bar', 'icon_set', 'top10', 'duplicate_values', 'unique_values', 'contains_text', 'time_period'], true)) {
            throw new MnbExcelException('Unsupported conditional-formatting type: ' . $type);
        }
        $this->nativeConditionalFormats[] = ['range' => $range, 'type' => $type] + $options;
        return $this;
    }

    public function conditionalCellIs(string $range, string $operator, mixed $formula, string|array $style): self
    {
        $formulas = is_array($formula) ? array_values($formula) : [$formula];
        return $this->conditionalFormatting($range, 'cell_is', [
            'operator' => $operator,
            'formulas' => $formulas,
            'style' => $style,
        ]);
    }

    public function conditionalExpression(string $range, string $formula, string|array $style): self
    {
        return $this->conditionalFormatting($range, 'expression', ['formulas' => [$formula], 'style' => $style]);
    }

    /** @param list<string> $colors */
    public function conditionalColorScale(string $range, array $colors = ['#F8696B', '#FFEB84', '#63BE7B']): self
    {
        return $this->conditionalFormatting($range, 'color_scale', ['colors' => array_values($colors)]);
    }

    public function conditionalDataBar(string $range, string $color = '#638EC6'): self
    {
        return $this->conditionalFormatting($range, 'data_bar', ['color' => $color]);
    }

    public function conditionalIconSet(string $range, string $iconSet = '3TrafficLights1'): self
    {
        return $this->conditionalFormatting($range, 'icon_set', ['icon_set' => $iconSet]);
    }

    /** Add native Excel data validation to a range. */
    public function dataValidation(string $range, string $type, array $options = []): self
    {
        $range = $this->normalizeRangeReference($range, 'data validation');
        $type = strtolower(trim($type));
        if (!in_array($type, ['list', 'whole', 'decimal', 'date', 'time', 'text_length', 'custom'], true)) {
            throw new MnbExcelException('Unsupported data-validation type: ' . $type);
        }
        $this->dataValidations[] = ['range' => $range, 'type' => $type] + $options;
        return $this;
    }

    /** @param list<string|int|float> $values */
    public function validationList(string $range, array $values, array $options = []): self
    {
        return $this->dataValidation($range, 'list', ['values' => array_values($values)] + $options);
    }

    /**
     * Add a native chart. Series items contain name, values and optional categories.
     *
     * @param list<array{name?:string,values:string,categories?:string}> $series
     * @param array<string,mixed> $options
     */
    public function addChart(string $type, string $title, array $series, array $options = []): self
    {
        $type = strtolower(trim($type));
        if (!in_array($type, ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'], true)) {
            throw new MnbExcelException('Unsupported chart type: ' . $type);
        }
        if ($series === []) {
            throw new MnbExcelException('A chart requires at least one series.');
        }
        foreach ($series as $item) {
            if (!isset($item['values']) || trim((string) $item['values']) === '') {
                throw new MnbExcelException('Every chart series requires a values range.');
            }
        }
        $sheetKey = $this->annotationSheetKey($options['sheet'] ?? null);
        $this->charts[$sheetKey][] = [
            'type' => $type,
            'title' => $title,
            'series' => array_values($series),
            'from' => strtoupper((string) ($options['from'] ?? $options['cell'] ?? 'A1')),
            'to' => strtoupper((string) ($options['to'] ?? 'H16')),
            'legend' => (string) ($options['legend'] ?? 'right'),
            'style' => max(1, min(48, (int) ($options['style'] ?? 10))),
            'vary_colors' => (bool) ($options['vary_colors'] ?? in_array($type, ['pie', 'doughnut'], true)),
        ];
        return $this;
    }

    /**
     * Build a reusable import template with header styling, instructions and validation rules.
     *
     * @param array<int|string,string|array<string,mixed>> $columns
     * @param array<string,mixed> $options
     */
    public static function importTemplate(array $columns, array $options = []): self
    {
        $headers = [];
        $definitions = [];
        foreach ($columns as $key => $definition) {
            if (is_array($definition)) {
                $name = (string) ($definition['header'] ?? $definition['name'] ?? $key);
                $definitions[] = $definition + ['header' => $name];
                $headers[] = $name;
            } else {
                $headers[] = (string) $definition;
                $definitions[] = ['header' => (string) $definition];
            }
        }
        $sampleRows = max(1, (int) ($options['sample_rows'] ?? 1));
        $rows = [];
        for ($i = 0; $i < $sampleRows; $i++) {
            $row = [];
            foreach ($definitions as $definition) {
                $row[] = $i === 0 ? ($definition['example'] ?? '') : '';
            }
            $rows[] = $row;
        }
        $builder = self::fromArray($rows)
            ->withHeader()
            ->columns($headers)
            ->styleHeader($options['header_style'] ?? 'mnb.header.blue')
            ->freezeHeader()
            ->autoFilter()
            ->autoWidth(['min' => 12, 'max' => 45, 'padding' => 3])
            ->metadata(['title' => (string) ($options['title'] ?? 'Import Template')]);

        if (isset($options['instructions']) && (string) $options['instructions'] !== '') {
            $builder->title((string) $options['instructions'], ['style' => 'mnb.subtitle', 'merge' => true]);
        }
        $lastRow = max(1000, (int) ($options['validation_rows'] ?? 10000));
        foreach ($definitions as $index => $definition) {
            $columnName = Coordinate::columnIndexToName($index + 1);
            $dataStart = ($options['instructions'] ?? '') !== '' ? 3 : 2;
            $range = $columnName . $dataStart . ':' . $columnName . $lastRow;
            if (isset($definition['list']) && is_array($definition['list'])) {
                $builder->validationList($range, $definition['list'], [
                    'allow_blank' => !($definition['required'] ?? false),
                    'prompt_title' => $definition['header'],
                    'prompt' => $definition['description'] ?? null,
                    'error' => $definition['error'] ?? null,
                ]);
            } elseif (isset($definition['validation']) && is_array($definition['validation'])) {
                $validation = $definition['validation'];
                $builder->dataValidation($range, (string) ($validation['type'] ?? 'custom'), $validation);
            }
            if (($definition['required'] ?? false) === true) {
                $headerRow = ($options['instructions'] ?? '') !== '' ? 2 : 1;
                $builder->comment($columnName . $headerRow, 'MNB PHPExcel', 'Required field' . (isset($definition['description']) ? ': ' . $definition['description'] : ''));
            }
        }
        return $builder;
    }

    /** @param array<string, mixed>|string $style */
    public function styleHeader(array|string $style): self
    {
        $this->headerStyle = $style;
        return $this;
    }

    public function escapeFormulaLikeText(bool $enabled = true): self
    {
        $this->escapeFormulaLikeText = $enabled;
        return $this;
    }

    /**
     * Merge one or more cell ranges. Example: mergeCells('A1:D1') or mergeCells(['A1:D1', 'A2:B2']).
     *
     * @param string|list<string> $ranges
     */
    public function mergeCells(string|array $ranges): self
    {
        foreach ((array) $ranges as $range) {
            $range = strtoupper(trim((string) $range));
            if (!preg_match('/^[A-Z]+\d+:[A-Z]+\d+$/', $range)) {
                throw new MnbExcelException('Invalid merge cell range: ' . $range);
            }
            $this->mergeCells[] = $range;
        }

        $this->mergeCells = array_values(array_unique($this->mergeCells));
        return $this;
    }

    public function columnWidth(int|string $column, float|int $width): self
    {
        if ($width <= 0) {
            throw new MnbExcelException('Column width must be greater than zero.');
        }

        $this->columnWidths[$column] = $width;
        return $this;
    }

    /** @param array<int|string, float|int> $widths */
    public function columnWidths(array $widths): self
    {
        foreach ($widths as $column => $width) {
            $this->columnWidth($column, $width);
        }
        return $this;
    }

    public function rowHeight(int $row, float|int $height): self
    {
        if ($row < 1) {
            throw new MnbExcelException('Row number must be greater than zero.');
        }
        if ($height <= 0) {
            throw new MnbExcelException('Row height must be greater than zero.');
        }

        $this->rowHeights[$row] = $height;
        return $this;
    }

    /** @param array<int, float|int> $heights */
    public function rowHeights(array $heights): self
    {
        foreach ($heights as $row => $height) {
            $this->rowHeight((int) $row, $height);
        }
        return $this;
    }

    /** @param array{sheet?:string,width?:int,height?:int,name?:string} $options */
    public function addImage(string $path, string $cell, array $options = []): self
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('Image file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }

        $cell = strtoupper(trim($cell));
        if (!preg_match('/^[A-Z]+\d+$/', $cell)) {
            throw new MnbExcelException('Invalid image cell reference: ' . $cell);
        }

        $image = ['path' => $path, 'cell' => $cell];
        if (isset($options['width'])) {
            $image['width'] = max(1, (int) $options['width']);
        }
        if (isset($options['height'])) {
            $image['height'] = max(1, (int) $options['height']);
        }
        if (isset($options['name'])) {
            $image['name'] = (string) $options['name'];
        }

        $sheetKey = $this->annotationSheetKey($options['sheet'] ?? null);
        $this->images[$sheetKey][] = $image;
        return $this;
    }

    /**
     * Add a clickable hyperlink to a generated XLSX worksheet.
     *
     * The link is written into worksheet XML and sheet relationships, then validated by
     * XlsxIntegrityValidator during save(). For multiple sheets, pass ['sheet' => 'Sheet name'].
     *
     * @param array{sheet?:string,tooltip?:string} $options
     */
    public function hyperlink(string $cell, string $url, ?string $display = null, array $options = []): self
    {
        $cell = $this->normalizeCellReference($cell, 'hyperlink');
        $url = trim($url);
        if ($url === '') {
            throw new MnbExcelException('Hyperlink URL cannot be empty.');
        }

        $sheetName = $this->annotationSheetKey($options['sheet'] ?? null);
        $link = [
            'cell' => $cell,
            'url' => $url,
        ];
        if ($display !== null && $display !== '') {
            $link['display'] = $display;
        }
        if (isset($options['tooltip']) && (string) $options['tooltip'] !== '') {
            $link['tooltip'] = (string) $options['tooltip'];
        }

        $this->hyperlinks[$sheetName][] = $link;
        return $this;
    }

    /**
     * Add many hyperlinks. Each item may contain: cell, url, display, tooltip, sheet.
     *
     * @param list<array{cell:string,url:string,display?:string,tooltip?:string,sheet?:string}> $hyperlinks
     */
    public function hyperlinks(array $hyperlinks): self
    {
        foreach ($hyperlinks as $link) {
            $options = [];
            if (isset($link['sheet'])) {
                $options['sheet'] = (string) $link['sheet'];
            }
            if (isset($link['tooltip'])) {
                $options['tooltip'] = (string) $link['tooltip'];
            }
            $this->hyperlink((string) $link['cell'], (string) $link['url'], isset($link['display']) ? (string) $link['display'] : null, $options);
        }

        return $this;
    }

    /**
     * Add a legacy Excel note/comment to a generated XLSX worksheet.
     *
     * Comments are written as xl/commentsN.xml plus a VML drawing part, matching the
     * structure used by Excel for classic notes. For multiple sheets, pass ['sheet' => 'Sheet name'].
     *
     * @param array{sheet?:string,width?:float|int,height?:float|int,visible?:bool} $options
     */
    public function comment(string $cell, string $author, string $text, array $options = []): self
    {
        $cell = $this->normalizeCellReference($cell, 'comment');
        $author = trim($author);
        $text = (string) $text;
        if ($author === '') {
            $author = 'MNB PHPExcel';
        }
        if ($text === '') {
            throw new MnbExcelException('Comment text cannot be empty.');
        }

        $sheetName = $this->annotationSheetKey($options['sheet'] ?? null);
        $comment = [
            'cell' => $cell,
            'author' => $author,
            'text' => $text,
        ];
        if (isset($options['width'])) {
            $comment['width'] = max(40, (float) $options['width']);
        }
        if (isset($options['height'])) {
            $comment['height'] = max(20, (float) $options['height']);
        }
        if (array_key_exists('visible', $options)) {
            $comment['visible'] = (bool) $options['visible'];
        }

        $this->comments[$sheetName][] = $comment;
        return $this;
    }

    /**
     * Add many comments. Each item may contain: cell, author, text, sheet, width, height, visible.
     *
     * @param list<array{cell:string,author?:string,text:string,sheet?:string,width?:float|int,height?:float|int,visible?:bool}> $comments
     */
    public function comments(array $comments): self
    {
        foreach ($comments as $comment) {
            $options = [];
            foreach (['sheet', 'width', 'height', 'visible'] as $key) {
                if (array_key_exists($key, $comment)) {
                    $options[$key] = $comment[$key];
                }
            }
            $this->comment(
                (string) $comment['cell'],
                isset($comment['author']) ? (string) $comment['author'] : 'MNB PHPExcel',
                (string) $comment['text'],
                $options
            );
        }

        return $this;
    }

    /**
     * Add a merged report title row before the header/data section.
     *
     * @param array<string,mixed> $options
     */
    public function title(string $title, array $options = []): self
    {
        $this->titleRows[] = [
            'values' => [$title],
            'style' => $options['style'] ?? 'mnb.title',
            'height' => $options['height'] ?? 28,
            'merge' => $options['merge'] ?? true,
        ];

        return $this;
    }

    /**
     * Add any custom row before the header/data section.
     *
     * @param string|list<mixed> $values
     * @param array<string,mixed> $options
     */
    public function titleRow(string|array $values, array $options = []): self
    {
        $this->titleRows[] = [
            'values' => $this->valuesToRow($values),
            'style' => $options['style'] ?? 'mnb.subtitle',
            'height' => $options['height'] ?? null,
            'merge' => $options['merge'] ?? false,
        ];

        return $this;
    }

    /**
     * Add summary rows after the data section.
     *
     * @param list<array<int|string,mixed>>|array<int|string,mixed> $rows
     * @param string|array<string,mixed>|null $style
     */
    public function summaryRows(array $rows, string|array|null $style = null): self
    {
        foreach ($this->normalizeReportRows($rows) as $row) {
            $this->summaryRows[] = [
                'values' => $row,
                'style' => $style ?? 'mnb.summary',
            ];
        }

        return $this;
    }

    /**
     * Add footer rows after summary rows.
     *
     * @param list<array<int|string,mixed>>|array<int|string,mixed>|list<string>|string $rows
     * @param string|array<string,mixed>|null $style
     */
    public function footerRows(array|string $rows, string|array|null $style = null): self
    {
        foreach ($this->normalizeReportRows($rows) as $row) {
            $this->footerRows[] = [
                'values' => $row,
                'style' => $style ?? 'mnb.footer',
                'merge' => count($row) === 1,
            ];
        }

        return $this;
    }

    /** @param array<string,mixed> $style */
    public function namedStyle(string $name, array $style): self
    {
        $name = trim($name);
        if ($name === '') {
            throw new MnbExcelException('Named style name cannot be empty.');
        }

        $this->namedStyles[$name] = $style;
        return $this;
    }

    /** @param string|array<string,mixed> $style */
    public function columnStyle(int|string $column, string|array $style): self
    {
        $this->columnStyles[$column] = $style;
        return $this;
    }

    /** @param array<int|string, string|array<string,mixed>> $styles */
    public function columnStyles(array $styles): self
    {
        foreach ($styles as $column => $style) {
            $this->columnStyle($column, $style);
        }
        return $this;
    }

    /** @param string|array<string,mixed> $style */
    public function rowStyle(int $row, string|array $style): self
    {
        if ($row < 1) {
            throw new MnbExcelException('Row number must be greater than zero.');
        }

        $this->rowStyles[$row] = $style;
        return $this;
    }

    /** @param array<int, string|array<string,mixed>> $styles */
    public function rowStyles(array $styles): self
    {
        foreach ($styles as $row => $style) {
            $this->rowStyle((int) $row, $style);
        }
        return $this;
    }

    /** @param string|array<string,mixed> $style */
    public function cellStyle(string $cell, string|array $style): self
    {
        $cell = strtoupper(trim($cell));
        if (!preg_match('/^[A-Z]+\d+$/', $cell)) {
            throw new MnbExcelException('Invalid cell reference: ' . $cell);
        }

        $this->cellStyles[$cell] = $style;
        return $this;
    }

    /** @param array<string, string|array<string,mixed>> $styles */
    public function cellStyles(array $styles): self
    {
        foreach ($styles as $cell => $style) {
            $this->cellStyle((string) $cell, $style);
        }
        return $this;
    }

    /** @param string|array<string,mixed> $style */
    public function rangeStyle(string $range, string|array $style): self
    {
        $range = strtoupper(trim($range));
        if (!preg_match('/^[A-Z]+\d+:[A-Z]+\d+$/', $range)) {
            throw new MnbExcelException('Invalid style range: ' . $range);
        }

        $this->rangeStyles[$range] = $style;
        return $this;
    }

    /** @param array<string, string|array<string,mixed>> $styles */
    public function rangeStyles(array $styles): self
    {
        foreach ($styles as $range => $style) {
            $this->rangeStyle((string) $range, $style);
        }
        return $this;
    }

    /** @param list<int|string> $columns */
    public function currencyColumns(array $columns, string $symbol = '$'): self
    {
        foreach ($columns as $column) {
            $this->columnStyle($column, [
                'format' => $symbol . '#,##0.00',
                'alignment' => ['horizontal' => 'right'],
            ]);
        }

        return $this;
    }

    /** @param list<int|string> $columns */
    public function percentageColumns(array $columns): self
    {
        foreach ($columns as $column) {
            $this->columnStyle($column, [
                'format' => '0.00%',
                'alignment' => ['horizontal' => 'right'],
            ]);
        }

        return $this;
    }

    /** @param list<int|string> $columns */
    public function dateStyleColumns(array $columns, string $format = 'yyyy-mm-dd'): self
    {
        foreach ($columns as $column) {
            $this->columnStyle($column, [
                'format' => $format,
                'alignment' => ['horizontal' => 'left'],
            ]);
        }

        return $this;
    }

    /** @param list<int|string> $columns */
    public function datetimeStyleColumns(array $columns, string $format = 'yyyy-mm-dd hh:mm:ss'): self
    {
        foreach ($columns as $column) {
            $this->columnStyle($column, [
                'format' => $format,
                'alignment' => ['horizontal' => 'left'],
            ]);
        }

        return $this;
    }

    /** @param list<int|string> $columns */
    public function integerColumns(array $columns): self
    {
        return $this->formatColumns($columns, 'integer', ['alignment' => ['horizontal' => 'right']]);
    }

    /** @param list<int|string> $columns */
    public function decimalColumns(array $columns, int $decimals = 2): self
    {
        $decimals = max(0, min(8, $decimals));
        $format = '#,##0' . ($decimals > 0 ? '.' . str_repeat('0', $decimals) : '');
        return $this->formatColumns($columns, $format, ['alignment' => ['horizontal' => 'right']]);
    }

    /** @param list<int|string> $columns */
    public function textStyleColumns(array $columns): self
    {
        return $this->formatColumns($columns, 'text', ['alignment' => ['horizontal' => 'left']]);
    }

    /**
     * Apply one number-format preset or custom Excel format to many columns.
     *
     * @param list<int|string> $columns
     * @param array<string,mixed> $extraStyle
     */
    public function formatColumns(array $columns, string $format, array $extraStyle = []): self
    {
        foreach ($columns as $column) {
            $this->columnStyle($column, array_replace_recursive(['format' => $format], $extraStyle));
        }

        return $this;
    }

    /**
     * Estimate column widths from visible row values. Manual columnWidth() values override estimates.
     *
     * @param bool|array<string,int> $enabled
     * @param array<string,int> $options
     */
    public function autoWidth(bool|array $enabled = true, array $options = []): self
    {
        if (is_array($enabled)) {
            $options = $enabled;
            $enabled = true;
        }

        $this->autoWidth = $enabled;
        foreach (['min', 'max', 'padding'] as $key) {
            if (isset($options[$key])) {
                $this->autoWidthOptions[$key] = max(0, (int) $options[$key]);
            }
        }
        if ($this->autoWidthOptions['max'] < $this->autoWidthOptions['min']) {
            $this->autoWidthOptions['max'] = $this->autoWidthOptions['min'];
        }

        return $this;
    }

    /**
     * Style data rows when a condition returns true.
     * Callback receives: original row, normalized row, data row number, Excel row number.
     *
     * @param callable(array<int|string,mixed>, list<mixed>, int, int):bool $condition
     * @param string|array<string,mixed> $style
     */
    public function conditionalRowStyle(callable $condition, string|array $style): self
    {
        $this->conditionalRowStyles[] = [
            'condition' => $condition,
            'style' => $style,
        ];

        return $this;
    }

    /** @param array<string,mixed> $metadata */
    public function metadata(array $metadata): self
    {
        foreach ($metadata as $key => $value) {
            if ($value === null) {
                unset($this->metadata[(string) $key]);
                continue;
            }
            $this->metadata[(string) $key] = $value;
        }

        return $this;
    }

    public function creator(string $creator): self
    {
        return $this->metadata(['creator' => $creator, 'last_modified_by' => $creator]);
    }

    public function company(string $company): self
    {
        return $this->metadata(['company' => $company]);
    }


    /**
     * Configure save-time XLSX integrity validation.
     *
     * Enabled by default for XLSX exports. When validation fails, the partial output file is deleted and
     * MnbExcelException is thrown before a corrupted workbook is returned to application code.
     *
     * Supported options include: require_xmlreader.
     *
     * @param array<string,mixed> $options
     */
    public function xlsxIntegrityValidation(bool $enabled = true, array $options = []): self
    {
        $this->xlsxIntegrityValidation = array_replace(['enabled' => $enabled], $options, ['enabled' => $enabled]);
        return $this;
    }

    /** Disable save-time XLSX integrity validation for special/debug workflows. */
    public function skipXlsxIntegrityValidation(): self
    {
        return $this->xlsxIntegrityValidation(false);
    }

    /** Require ext-xmlreader for strict XML well-formedness checks during XLSX save validation. */
    public function strictXlsxIntegrityValidation(bool $enabled = true): self
    {
        return $this->xlsxIntegrityValidation(true, ['require_xmlreader' => $enabled]);
    }

    /**
     * Preserve unsupported/advanced XLSX package objects from a source workbook while writing new row data.
     * This keeps small-file output practical without silently dropping comments, drawings, tables, legacy notes,
     * hyperlinks, VML, embedded media, pivot/chart/table package parts, and related worksheet relationships.
     *
     * The safest use is a same-sheet-order template: read/prepare rows, then save into a new file using this
     * method with the original XLSX path.
     *
     * @param array<string,mixed> $options
     */
    public function preserveAdvancedObjectsFrom(string $xlsxPath, array $options = []): self
    {
        $realPath = realpath($xlsxPath);
        if ($realPath === false || !is_file($realPath)) {
            throw MnbExcelException::withCode('Advanced-object template XLSX not found: ' . $xlsxPath, ErrorCode::FILE_NOT_FOUND, ['path' => $xlsxPath]);
        }

        $this->advancedObjectPreservation = array_replace([
            'path' => $realPath,
            'mode' => 'advanced_objects',
            'sheet_match' => 'index',
            'preserve_sheet_relationships' => true,
            'preserve_sheet_elements' => true,
            'copy_package_parts' => true,
        ], $options);

        return $this;
    }

    /** Alias for preserveAdvancedObjectsFrom(). */
    public function preserveAdvancedExcelObjects(string $xlsxPath, array $options = []): self
    {
        return $this->preserveAdvancedObjectsFrom($xlsxPath, $options);
    }


    /**
     * Preserve working pivot tables from an XLSX template and bind their cache
     * to a new source range. Excel refreshes the pivot cache when the file opens.
     *
     * This is template-driven pivot support: create/layout the pivot once in
     * Excel, then use this method to replace its source data safely.
     */
    public function preservePivotTablesFrom(
        string $xlsxPath,
        string $sourceSheet,
        string $sourceRange,
        array $options = []
    ): self {
        $sourceSheet = trim($sourceSheet);
        if ($sourceSheet === '') {
            throw new MnbExcelException('Pivot source sheet cannot be empty.');
        }
        $sourceRange = $this->normalizeRangeReference($sourceRange, 'pivot source');
        return $this->preserveAdvancedObjectsFrom($xlsxPath, array_replace([
            'pivot_source_sheet' => $sourceSheet,
            'pivot_source_range' => $sourceRange,
            'pivot_refresh_on_load' => true,
            'preserve_workbook_pivot_caches' => true,
        ], $options));
    }

    /**
     * Add a dedicated import summary sheet, useful beside failed-row reports.
     *
     * @param array<string,mixed> $summary
     * @param array<string,mixed> $validationResult
     */
    public function withImportSummarySheet(array $summary, array $validationResult = [], string $sheetName = 'Import Summary'): self
    {
        $this->importSummarySheet = [
            'name' => $this->sanitizeSheetName($sheetName),
            'rows' => $this->buildImportSummaryRows($summary, $validationResult),
            'style' => 'mnb.header.blue',
        ];

        return $this;
    }

    /**
     * Configure CSV read/write behavior for CSV exports generated by this builder.
     *
     * @param array<string,mixed> $options
     */
    public function csvOptions(array $options): self
    {
        $this->csvOptions = array_replace($this->csvOptions, $options);
        return $this;
    }

    public function csvDialect(string $dialect): self
    {
        return $this->csvOptions(['dialect' => $dialect]);
    }

    public function csvDelimiter(string $delimiter): self
    {
        return $this->csvOptions(['delimiter' => $delimiter]);
    }

    public function csvEnclosure(string $enclosure): self
    {
        return $this->csvOptions(['enclosure' => $enclosure]);
    }

    public function csvEscape(string $escape): self
    {
        return $this->csvOptions(['escape' => $escape]);
    }

    public function csvBom(bool $enabled = true): self
    {
        return $this->csvOptions(['bom' => $enabled]);
    }

    public function csvEncoding(string $encoding): self
    {
        return $this->csvOptions(['encoding' => $encoding]);
    }

    public function csvInjectionPolicy(string $policy): self
    {
        return $this->csvOptions(['injection_policy' => $policy]);
    }

    /**
     * Configure cell text safety behavior for XLSX/CSV export.
     *
     * Supported options: max_text_length, long_text_policy, control_char_policy,
     * large_number_as_text, and allow_precision_loss.
     *
     * @param array<string,mixed> $options
     */
    public function cellSafety(array $options): self
    {
        $this->cellSafetyOptions = array_replace($this->cellSafetyOptions, $options);
        return $this;
    }

    public function maxCellTextLength(int $length, string $policy = 'truncate'): self
    {
        if ($length < 1) {
            throw new MnbExcelException('Maximum cell text length must be greater than zero.');
        }

        return $this->cellSafety([
            'max_text_length' => $length,
            'long_text_policy' => $policy,
        ]);
    }

    public function controlCharPolicy(string $policy): self
    {
        return $this->cellSafety(['control_char_policy' => $policy]);
    }

    /**
     * Configure explicit formula handling.
     * Policies: safe, allow, block. Default: safe.
     *
     * @param array<string,mixed>|string $policyOrOptions
     */
    public function formulaPolicy(array|string $policyOrOptions): self
    {
        $options = is_array($policyOrOptions) ? $policyOrOptions : ['formula_policy' => $policyOrOptions];
        $this->formulaOptions = array_replace($this->formulaOptions, $options);
        return $this;
    }

    public function allowFormulas(bool $enabled = true): self
    {
        return $this->formulaPolicy($enabled ? 'allow' : 'block');
    }

    /**
     * Scan source rows for formula-like text, unsafe explicit formulas, invalid XML characters,
     * overlong text, and long numeric text that may lose precision in Excel.
     *
     * @return array{status:string,total_issues:int,issues:list<array<string,mixed>>}
     */
    public function scanCellSafety(): array
    {
        $rows = [];
        foreach ($this->sourceSheets as $sheetRows) {
            foreach ($sheetRows as $row) {
                if (is_array($row)) {
                    $rows[] = $row;
                }
            }
        }

        return (new CellSafetyScanner())->scan($rows, array_replace($this->cellSafetyOptions, $this->formulaOptions));
    }

    public function saveSafe(string $directory, string $filename, string $extension = 'xlsx'): string
    {
        $extension = strtolower(ltrim($extension, '.'));
        if (!in_array($extension, ['xlsx', 'csv', 'json', 'xml'], true)) {
            throw MnbExcelException::withCode('Unsupported safe save extension: ' . $extension, ErrorCode::UNSUPPORTED_FORMAT, ['extension' => $extension]);
        }

        if (!is_dir($directory) && !mkdir($directory, 0775, true) && !is_dir($directory)) {
            throw MnbExcelException::withCode('Unable to create directory: ' . $directory, ErrorCode::DIRECTORY_CREATE_FAILED, ['directory' => $directory]);
        }

        $safe = self::safeFileName($filename, $extension);
        return $this->save(rtrim($directory, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . $safe);
    }

    public static function safeFileName(string $filename, string $extension = ''): string
    {
        $extension = strtolower(ltrim(trim($extension), '.'));
        $filename = trim($filename);
        $existingExtension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        if ($extension === '' && $existingExtension !== '') {
            $extension = $existingExtension;
        }
        $base = $existingExtension !== '' ? substr($filename, 0, -(strlen($existingExtension) + 1)) : $filename;
        $base = preg_replace('/[^A-Za-z0-9._-]+/', '_', $base) ?: 'export';
        $base = trim($base, '._-');
        $base = $base === '' ? 'export' : substr($base, 0, 120);

        return $extension !== '' ? $base . '.' . $extension : $base;
    }

    private function normalizeCellReference(string $cell, string $purpose): string
    {
        $cell = strtoupper(trim($cell));
        if (!preg_match('/^[A-Z]{1,3}[1-9][0-9]*$/', $cell)) {
            throw new MnbExcelException('Invalid ' . $purpose . ' cell reference: ' . $cell);
        }

        return $cell;
    }

    private function annotationSheetKey(mixed $sheetName = null): string
    {
        if ($sheetName === null || trim((string) $sheetName) === '') {
            return (string) (array_key_first($this->sourceSheets) ?? 'Sheet1');
        }

        $requested = (string) $sheetName;
        if (array_key_exists($requested, $this->sourceSheets)) {
            return $requested;
        }

        $sanitized = SheetNameAllocator::sanitize($requested);
        $matches = [];
        foreach (array_keys($this->sourceSheets) as $sourceSheetName) {
            if (SheetNameAllocator::sanitize((string) $sourceSheetName) === $sanitized) {
                $matches[] = (string) $sourceSheetName;
            }
        }

        if (count($matches) === 1) {
            return $matches[0];
        }
        if (count($matches) > 1) {
            throw new MnbExcelException('Annotation sheet name is ambiguous; use the exact source sheet name: ' . $requested);
        }

        throw new MnbExcelException('Annotation sheet not found: ' . $requested);
    }

    /** @return list<array{path:string,cell:string,width?:int,height?:int,name?:string}> */
    private function imagesForSheet(string $sheetName): array
    {
        return $this->images[$sheetName] ?? [];
    }

    /** @return list<array{cell:string,url:string,display?:string,tooltip?:string}> */
    private function hyperlinksForSheet(string $sheetName): array
    {
        return $this->hyperlinks[$sheetName] ?? [];
    }

    /** @return list<array{cell:string,author:string,text:string,width?:float|int,height?:float|int,visible?:bool}> */
    private function commentsForSheet(string $sheetName): array
    {
        return $this->comments[$sheetName] ?? [];
    }

    public function reportTemplate(string $name): self
    {
        $name = strtolower(trim($name));
        if (!in_array($name, ['simple', 'business', 'finance'], true)) {
            throw new MnbExcelException('Unknown report template: ' . $name);
        }

        if ($name === 'simple') {
            $this->styleHeader('mnb.header');
            return $this;
        }

        if ($name === 'business') {
            $this->styleHeader('mnb.header.blue');
            $this->rowHeight(1, 28);
            return $this;
        }

        $this->styleHeader('mnb.header.green');
        return $this;
    }

    public function save(string $path): string
    {
        $extension = strtolower(pathinfo($path, PATHINFO_EXTENSION));
        $workbook = $this->toWorkbookData();

        if ($extension === 'csv') {
            if (count($workbook->sheets) > 1) {
                throw MnbExcelException::withCode('CSV supports only one sheet. Use saveCsvSheets() for multiple CSV files.', ErrorCode::UNSUPPORTED_FORMAT, ['format' => 'csv']);
            }

            (new CsvWriter())->write($workbook->sheets[0], $path, $this->csvOptions);
            return $path;
        }

        if ($extension === 'xlsx') {
            (new XlsxWriter())->write($workbook, $path);
            return $path;
        }

        if ($extension === 'json') {
            (new JsonWriter())->writeWorkbook($workbook, $path);
            return $path;
        }

        if ($extension === 'xml') {
            (new XmlWriter())->writeWorkbook($workbook, $path);
            return $path;
        }

        throw MnbExcelException::withCode('Unsupported output format: ' . $extension, ErrorCode::UNSUPPORTED_FORMAT, ['extension' => $extension]);
    }

    /** @param array<string,mixed> $options */
    public function saveJson(string $path, array $options = []): string
    {
        (new JsonWriter())->writeWorkbook($this->toWorkbookData(), $path, $options);
        return $path;
    }

    /** @param array<string,mixed> $options */
    public function toJson(array $options = []): string
    {
        return (new JsonWriter())->workbookToString($this->toWorkbookData(), $options);
    }

    /** @param array<string,mixed> $options */
    public function saveXml(string $path, array $options = []): string
    {
        (new XmlWriter())->writeWorkbook($this->toWorkbookData(), $path, $options);
        return $path;
    }

    /** @param array<string,mixed> $options */
    public function toXml(array $options = []): string
    {
        return (new XmlWriter())->workbookToString($this->toWorkbookData(), $options);
    }

    /** @return list<string> */
    public function saveCsvSheets(string $directory): array
    {
        if (!is_dir($directory) && !mkdir($directory, 0775, true) && !is_dir($directory)) {
            throw MnbExcelException::withCode('Unable to create directory: ' . $directory, ErrorCode::DIRECTORY_CREATE_FAILED, ['directory' => $directory]);
        }

        $files = [];
        foreach ($this->toWorkbookData()->sheets as $sheet) {
            $file = rtrim($directory, DIRECTORY_SEPARATOR) . DIRECTORY_SEPARATOR . preg_replace('/[^A-Za-z0-9_-]+/', '_', $sheet->name) . '.csv';
            (new CsvWriter())->write($sheet, $file, $this->csvOptions);
            $files[] = $file;
        }

        return $files;
    }

    /** @return list<list<mixed>> */
    public function toArray(): array
    {
        return $this->toWorkbookData()->sheets[0]->rows;
    }

    /** @param array<string,mixed> $options */
    public function importToSql(PDO|array|string|null $pdo, string $table, array $options = []): array
    {
        $connection = DatabaseConnectionFactory::make($pdo, is_array($options['db'] ?? null) ? $options['db'] : []);
        return (new SqlImporter())->importRows($connection, $table, $this->rowsForSqlImport(), $options);
    }

    /**
     * Keep SQL imports associative when source rows are associative.
     *
     * @return list<array<string,mixed>|list<mixed>>
     */
    private function rowsForSqlImport(): array
    {
        $sheetRows = reset($this->sourceSheets);
        if (!is_array($sheetRows)) {
            return [];
        }

        $rows = [];
        foreach ($sheetRows as $row) {
            if (!is_array($row)) {
                throw new MnbExcelException('Every input row must be an array.');
            }

            /** @var array<int|string,mixed> $row */
            if (Arr::isAssoc($row)) {
                $normalized = [];
                foreach ($row as $key => $value) {
                    $normalized[(string) $key] = $this->normalizeValue($value, (string) $key);
                }
                $rows[] = $normalized;
            } else {
                $normalized = [];
                foreach (array_values($row) as $index => $value) {
                    $normalized[] = $this->normalizeValue($value, (string) $index);
                }
                $rows[] = $normalized;
            }
        }

        return $rows;
    }

    public function toWorkbookData(): WorkbookData
    {
        $sheets = [];
        $sheetNameAllocator = new SheetNameAllocator();
        foreach ($this->sourceSheets as $sheetName => $rows) {
            $body = $this->normalizeRows($rows);
            $headerRowIndex = null;
            $finalRows = [];
            $rowStyles = $this->rowStyles;
            $rowHeights = $this->rowHeights;
            $mergeCells = $this->mergeCells;

            foreach ($this->titleRows as $sectionRow) {
                $finalRows[] = $sectionRow['values'];
                $rowNumber = count($finalRows);
                if (isset($sectionRow['style'])) {
                    $rowStyles[$rowNumber] = $sectionRow['style'];
                }
                if (isset($sectionRow['height']) && $sectionRow['height'] !== null) {
                    $rowHeights[$rowNumber] = $sectionRow['height'];
                }
            }

            if ($this->withHeader && $body !== []) {
                $headerRowIndex = count($finalRows) + 1;
            }

            $sourceRowsList = array_values($rows);
            foreach ($body as $bodyIndex => $bodyRow) {
                $finalRows[] = $bodyRow;
                $excelRowNumber = count($finalRows);
                $sourceIndex = $bodyIndex - ($this->withHeader ? 1 : 0);
                if ($sourceIndex >= 0 && isset($sourceRowsList[$sourceIndex]) && is_array($sourceRowsList[$sourceIndex])) {
                    $conditionalStyle = $this->styleForConditionalRow($sourceRowsList[$sourceIndex], $bodyRow, $sourceIndex + 1, $excelRowNumber);
                    if ($conditionalStyle !== null) {
                        $rowStyles[$excelRowNumber] = $conditionalStyle;
                    }
                }
            }

            foreach ($this->summaryRows as $sectionRow) {
                $finalRows[] = $sectionRow['values'];
                $rowNumber = count($finalRows);
                if (isset($sectionRow['style'])) {
                    $rowStyles[$rowNumber] = $sectionRow['style'];
                }
                if (isset($sectionRow['height']) && $sectionRow['height'] !== null) {
                    $rowHeights[$rowNumber] = $sectionRow['height'];
                }
            }

            foreach ($this->footerRows as $sectionRow) {
                $finalRows[] = $sectionRow['values'];
                $rowNumber = count($finalRows);
                if (isset($sectionRow['style'])) {
                    $rowStyles[$rowNumber] = $sectionRow['style'];
                }
                if (isset($sectionRow['height']) && $sectionRow['height'] !== null) {
                    $rowHeights[$rowNumber] = $sectionRow['height'];
                }
            }

            $maxColumns = $this->maxColumnCount($finalRows);
            foreach ($this->titleRows as $index => $sectionRow) {
                $rowNumber = $index + 1;
                if (($sectionRow['merge'] ?? false) && $maxColumns > 1) {
                    $mergeCells[] = 'A' . $rowNumber . ':' . Coordinate::columnIndexToName($maxColumns) . $rowNumber;
                }
            }

            $footerStart = count($this->titleRows) + count($body) + count($this->summaryRows) + 1;
            foreach ($this->footerRows as $index => $sectionRow) {
                $rowNumber = $footerStart + $index;
                if (($sectionRow['merge'] ?? false) && $maxColumns > 1) {
                    $mergeCells[] = 'A' . $rowNumber . ':' . Coordinate::columnIndexToName($maxColumns) . $rowNumber;
                }
            }

            /** @var array<int|string,mixed> $first */
            $first = is_array(reset($rows)) ? reset($rows) : [];
            $keys = $this->columnKeys($first);

            $sheets[] = new WorksheetData(
                name: $sheetNameAllocator->allocate((string) $sheetName),
                rows: $finalRows,
                columns: $this->columns,
                hasHeader: $this->withHeader,
                textColumns: $this->textColumns,
                dateColumns: $this->dateColumns,
                numberColumns: $this->numberColumns,
                freezeHeader: $this->freezeHeader,
                autoFilter: $this->autoFilter,
                escapeFormulaLikeText: $this->escapeFormulaLikeText,
                headerStyle: $this->headerStyle,
                mergeCells: array_values(array_unique($mergeCells)),
                columnWidths: $this->resolvedColumnWidths($finalRows),
                rowHeights: $rowHeights,
                images: $this->imagesForSheet((string) $sheetName),
                headerRowIndex: $headerRowIndex,
                namedStyles: $this->namedStyles,
                columnStyles: $this->resolveColumnStyleIndexes($this->effectiveColumnStyles(), $keys),
                rowStyles: $rowStyles,
                cellStyles: $this->cellStyles,
                rangeStyles: $this->rangeStyles,
                hyperlinks: $this->hyperlinksForSheet((string) $sheetName),
                comments: $this->commentsForSheet((string) $sheetName),
                sourceColumnKeys: $keys,
                dataRowStart: count($this->titleRows) + ($this->withHeader ? 1 : 0),
                dataRowCount: count($rows),
                freezeRows: $this->freezeRows,
                freezeColumns: $this->freezeColumns,
                freezeTopLeftCell: $this->freezeTopLeftCell,
                autoFilterRange: $this->autoFilterRange,
                filterColumns: $this->filterColumns,
                conditionalFormats: $this->nativeConditionalFormats,
                dataValidations: $this->dataValidations,
                charts: $this->chartsForSheet((string) $sheetName)
            );
        }

        if ($this->importSummarySheet !== null) {
            $summaryRows = $this->importSummarySheet['rows'];
            $summaryHeaderRow = 1;
            $sheets[] = new WorksheetData(
                name: $sheetNameAllocator->allocate((string) $this->importSummarySheet['name']),
                rows: $summaryRows,
                hasHeader: true,
                freezeHeader: true,
                autoFilter: true,
                headerStyle: $this->importSummarySheet['style'] ?? 'mnb.header.blue',
                columnWidths: $this->resolvedColumnWidths($summaryRows),
                rowHeights: [1 => 24],
                headerRowIndex: $summaryHeaderRow,
                namedStyles: $this->namedStyles,
            );
        }

        $metadata = $this->metadata;
        if ($this->advancedObjectPreservation !== []) {
            $metadata['_mnb_preserve_xlsx_package'] = $this->advancedObjectPreservation;
        }
        $metadata['_mnb_xlsx_integrity_validation'] = $this->xlsxIntegrityValidation;

        return new WorkbookData($sheets, $metadata);
    }

    /**
     * @param array<int|string, mixed> $sourceRows
     * @return list<list<mixed>>
     */
    private function normalizeRows(array $sourceRows): array
    {
        if ($sourceRows === []) {
            return [];
        }

        $first = reset($sourceRows);
        if (!is_array($first)) {
            throw new MnbExcelException('Input array must be a list of row arrays.');
        }

        /** @var array<int|string, mixed> $first */
        $isAssoc = Arr::isAssoc($first);
        $keys = $this->columnKeys($first);
        $rows = [];

        if ($this->withHeader) {
            if ($this->columns !== []) {
                $rows[] = array_values($this->columns);
            } elseif ($isAssoc) {
                $rows[] = Arr::stringKeys($first);
            }
        }

        foreach ($sourceRows as $row) {
            if (!is_array($row)) {
                throw new MnbExcelException('Every input row must be an array.');
            }

            /** @var array<int|string, mixed> $row */
            if ($isAssoc) {
                $normalized = [];
                foreach ($keys as $key) {
                    $normalized[] = $this->normalizeValue($row[$key] ?? null, (string) $key);
                }
                $rows[] = $normalized;
                continue;
            }

            $normalized = [];
            foreach (array_values($row) as $i => $value) {
                $normalized[] = $this->normalizeValue($value, (string) $i);
            }
            $rows[] = $normalized;
        }

        return $rows;
    }

    /** @param array<int|string,mixed> $first */
    private function columnKeys(array $first): array
    {
        if ($this->columns !== []) {
            return array_keys($this->columns);
        }

        return Arr::isAssoc($first) ? array_keys($first) : [];
    }

    private function normalizeValue(mixed $value, string $columnKey): mixed
    {
        if (is_array($value) && array_key_exists('type', $value) && (array_key_exists('value', $value) || array_key_exists('formula', $value))) {
            $value = CellValue::fromArray($value);
        }

        if ($value instanceof CellValue) {
            return $this->normalizeCellValue($value);
        }

        if ($this->escapeFormulaLikeText) {
            $value = ValueSanitizer::escapeFormulaLikeText($value);
        }

        if ($value instanceof \DateTimeInterface) {
            $format = $this->dateColumns[$columnKey] ?? 'Y-m-d H:i:s';
            return $this->sanitizeText($value->format($format));
        }

        if (array_key_exists($columnKey, $this->dateColumns) && is_string($value) && $value !== '') {
            $timestamp = strtotime($value);
            if ($timestamp !== false) {
                return $this->sanitizeText(date($this->dateColumns[$columnKey], $timestamp));
            }
        }

        if (in_array($columnKey, $this->textColumns, true) && $value !== null) {
            return $this->sanitizeText((string) $value);
        }

        if (in_array($columnKey, $this->numberColumns, true) && is_numeric($value)) {
            if (is_string($value) && ($this->cellSafetyOptions['large_number_as_text'] ?? true) && !($this->cellSafetyOptions['allow_precision_loss'] ?? false) && ValueSanitizer::isLargeIntegerString($value)) {
                return $this->sanitizeText($value);
            }
            return $value + 0;
        }

        $normalized = ValueSanitizer::normalizeScalar($value);
        return is_string($normalized) ? $this->sanitizeText($normalized) : $normalized;
    }

    private function normalizeCellValue(CellValue $cell): CellValue
    {
        if ($cell->type() === CellValue::TYPE_FORMULA) {
            FormulaGuard::assertSafe((string) $cell->value(), $this->formulaOptions);
            return $cell;
        }

        if ($cell->type() === CellValue::TYPE_TEXT) {
            return CellValue::text($this->sanitizeText((string) $cell->value()));
        }

        if ($cell->type() === CellValue::TYPE_DATE) {
            $value = $cell->value();
            if ($value instanceof \DateTimeInterface) {
                return CellValue::date($value, $cell->options());
            }
            return CellValue::date($this->sanitizeText((string) $value), $cell->options());
        }

        return $cell;
    }

    private function sanitizeText(string $value): string
    {
        return ValueSanitizer::sanitizeCellText($value, $this->cellSafetyOptions);
    }

    private function normalizeRangeReference(string $range, string $context): string
    {
        $range = strtoupper(trim($range));
        if (preg_match('/^[A-Z]+\d+:[A-Z]+\d+$/', $range) !== 1) {
            throw new MnbExcelException('Invalid ' . $context . ' range: ' . $range);
        }
        return $range;
    }

    /** @return list<array<string,mixed>> */
    private function chartsForSheet(string $sheetName): array
    {
        return array_merge($this->charts['*'] ?? [], $this->charts[$sheetName] ?? []);
    }

    private function sanitizeSheetName(string $name): string
    {
        return SheetNameAllocator::sanitize($name);
    }

    /** @param string|list<mixed> $values */
    private function valuesToRow(string|array $values): array
    {
        return is_array($values) ? array_values($values) : [$values];
    }

    /**
     * @param mixed $rows
     * @return list<list<mixed>>
     */
    private function normalizeReportRows(mixed $rows): array
    {
        if (is_string($rows)) {
            return [[$rows]];
        }

        if (!is_array($rows)) {
            throw new MnbExcelException('Report rows must be string or array.');
        }

        if ($rows === []) {
            return [];
        }

        $first = reset($rows);
        if (!is_array($first)) {
            return [array_values($rows)];
        }

        $normalized = [];
        foreach ($rows as $row) {
            if (!is_array($row)) {
                throw new MnbExcelException('Every report row must be an array.');
            }
            $normalized[] = array_values($row);
        }

        return $normalized;
    }

    /** @param list<list<mixed>> $rows */
    private function maxColumnCount(array $rows): int
    {
        $max = 0;
        foreach ($rows as $row) {
            $max = max($max, count($row));
        }
        return $max;
    }

    /** @return array<int|string, string|array<string,mixed>> */
    private function effectiveColumnStyles(): array
    {
        $styles = [];

        foreach ($this->textColumns as $column) {
            $styles[$column] = $this->mergeStyleDefinitions($styles[$column] ?? [], ['format' => 'text']);
        }
        foreach ($this->dateColumns as $column => $_format) {
            $styles[$column] = $this->mergeStyleDefinitions($styles[$column] ?? [], ['format' => 'date']);
        }
        foreach ($this->numberColumns as $column) {
            $styles[$column] = $this->mergeStyleDefinitions($styles[$column] ?? [], ['format' => 'number']);
        }

        foreach ($this->columnStyles as $column => $style) {
            $styles[$column] = $this->mergeStyleDefinitions($styles[$column] ?? [], $style);
        }

        return $styles;
    }

    /** @param string|array<string,mixed>|mixed $base @param string|array<string,mixed>|mixed $override */
    private function mergeStyleDefinitions(mixed $base, mixed $override): mixed
    {
        if (is_string($override)) {
            return $override;
        }
        if (!is_array($override)) {
            return $base;
        }
        if (!is_array($base)) {
            return $override;
        }

        return array_replace_recursive($base, $override);
    }

    /**
     * @param array<int|string, string|array<string,mixed>> $styles
     * @param list<int|string> $keys
     * @return array<int, string|array<string,mixed>>
     */
    private function resolveColumnStyleIndexes(array $styles, array $keys): array
    {
        $resolved = [];
        foreach ($styles as $column => $style) {
            if (is_int($column) || ctype_digit((string) $column)) {
                $index = (int) $column;
            } else {
                $match = array_search($column, $keys, true);
                if ($match !== false) {
                    $index = (int) $match + 1;
                } elseif (preg_match('/^[A-Z]+$/i', (string) $column)) {
                    $index = Coordinate::columnNameToIndex((string) $column);
                } else {
                    continue;
                }
            }

            if ($index > 0) {
                $resolved[$index] = $style;
            }
        }

        return $resolved;
    }

    /**
     * @param array<int|string,mixed> $sourceRow
     * @param list<mixed> $normalizedRow
     */
    private function styleForConditionalRow(array $sourceRow, array $normalizedRow, int $dataRowNumber, int $excelRowNumber): string|array|null
    {
        $matchedStyle = null;
        foreach ($this->conditionalRowStyles as $rule) {
            $condition = $rule['condition'];
            if ((bool) $condition($sourceRow, $normalizedRow, $dataRowNumber, $excelRowNumber)) {
                $matchedStyle = $rule['style'];
            }
        }

        return $matchedStyle;
    }

    /**
     * @param list<list<mixed>> $rows
     * @return array<int,float|int>
     */
    private function resolvedColumnWidths(array $rows): array
    {
        if (!$this->autoWidth) {
            return $this->columnWidths;
        }

        $widths = $this->estimateColumnWidths($rows);
        foreach ($this->columnWidths as $column => $width) {
            $index = is_int($column) || ctype_digit((string) $column)
                ? (int) $column
                : Coordinate::columnNameToIndex((string) $column);
            if ($index > 0) {
                $widths[$index] = $width;
            }
        }

        return $widths;
    }

    /**
     * @param list<list<mixed>> $rows
     * @return array<int,int>
     */
    private function estimateColumnWidths(array $rows): array
    {
        $widths = [];
        $min = max(1, $this->autoWidthOptions['min']);
        $max = max($min, $this->autoWidthOptions['max']);
        $padding = max(0, $this->autoWidthOptions['padding']);

        foreach ($rows as $row) {
            foreach (array_values($row) as $index => $value) {
                $column = $index + 1;
                $text = $this->valueForWidth($value);
                $length = function_exists('mb_strlen') ? mb_strlen($text, 'UTF-8') : strlen($text);
                $candidate = min($max, max($min, $length + $padding));
                $widths[$column] = max($widths[$column] ?? $min, $candidate);
            }
        }

        return $widths;
    }

    private function valueForWidth(mixed $value): string
    {
        if ($value === null) {
            return '';
        }
        if ($value instanceof \DateTimeInterface) {
            return $value->format('Y-m-d H:i:s');
        }
        if (is_bool($value)) {
            return $value ? 'TRUE' : 'FALSE';
        }

        $text = (string) $value;
        $parts = preg_split('/\R/u', $text) ?: [$text];
        usort($parts, static fn (string $a, string $b): int => strlen($b) <=> strlen($a));
        return $parts[0] ?? '';
    }

    /**
     * @param array<string,mixed> $summary
     * @param array<string,mixed> $validationResult
     * @return list<array<int,mixed>>
     */
    private function buildImportSummaryRows(array $summary, array $validationResult = []): array
    {
        $rows = [['Metric', 'Value']];
        $flat = $this->flattenSummary($summary);
        foreach ($flat as $metric => $value) {
            $rows[] = [$metric, $this->summaryValueToString($value)];
        }

        if ($validationResult !== []) {
            if (isset($validationResult['valid']) && is_array($validationResult['valid'])) {
                $rows[] = ['valid_rows', count($validationResult['valid'])];
            }
            if (isset($validationResult['failed']) && is_array($validationResult['failed'])) {
                $rows[] = ['failed_rows', count($validationResult['failed'])];
            }
        }

        return $rows;
    }

    /** @return array<string,mixed> */
    private function flattenSummary(array $summary, string $prefix = ''): array
    {
        $flat = [];
        foreach ($summary as $key => $value) {
            $name = $prefix === '' ? (string) $key : $prefix . '.' . (string) $key;
            if (is_array($value) && $this->isListOfScalars($value)) {
                $flat[$name] = implode(', ', array_map(static fn (mixed $item): string => (string) $item, $value));
                continue;
            }
            if (is_array($value) && count($value) <= 20) {
                foreach ($this->flattenSummary($value, $name) as $childKey => $childValue) {
                    $flat[$childKey] = $childValue;
                }
                continue;
            }
            if (is_array($value)) {
                $flat[$name] = count($value);
                continue;
            }
            $flat[$name] = $value;
        }

        return $flat;
    }

    private function isListOfScalars(array $value): bool
    {
        if ($value === []) {
            return true;
        }
        foreach ($value as $item) {
            if (is_array($item) || is_object($item)) {
                return false;
            }
        }
        return array_is_list($value);
    }

    private function summaryValueToString(mixed $value): string|int|float|bool|null
    {
        if ($value === null || is_int($value) || is_float($value) || is_bool($value) || is_string($value)) {
            return $value;
        }
        if ($value instanceof \DateTimeInterface) {
            return $value->format('Y-m-d H:i:s');
        }

        $json = json_encode($value, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
        return $json === false ? (string) $value : $json;
    }

    private function registerDefaultReportStyles(): void
    {
        $this->namedStyles = [
            'mnb.title' => [
                'font' => ['bold' => true, 'size' => 18, 'color' => '#0B1220'],
                'fill' => '#F8FAFC',
                'alignment' => ['horizontal' => 'center', 'vertical' => 'center'],
                'border' => ['color' => '#CBD5E1'],
            ],
            'mnb.subtitle' => [
                'font' => ['italic' => true, 'size' => 11, 'color' => '#475569'],
                'fill' => '#F8FAFC',
                'alignment' => ['horizontal' => 'center', 'vertical' => 'center'],
            ],
            'mnb.header' => [
                'font' => ['bold' => true, 'size' => 12, 'color' => '#111827'],
                'fill' => '#EEF4FF',
                'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap_text' => true],
                'border' => ['color' => '#D0D7DE'],
            ],
            'mnb.header.blue' => [
                'font' => ['bold' => true, 'size' => 12, 'color' => '#FFFFFF'],
                'fill' => '#1F6FEB',
                'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap_text' => true],
                'border' => ['color' => '#BBD6FF'],
            ],
            'mnb.header.green' => [
                'font' => ['bold' => true, 'size' => 12, 'color' => '#FFFFFF'],
                'fill' => '#0F766E',
                'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap_text' => true],
                'border' => ['color' => '#99F6E4'],
            ],
            'mnb.summary' => [
                'font' => ['bold' => true, 'color' => '#111827'],
                'fill' => '#ECFDF5',
                'alignment' => ['horizontal' => 'right', 'vertical' => 'center'],
                'border' => ['color' => '#A7F3D0'],
            ],
            'mnb.footer' => [
                'font' => ['italic' => true, 'size' => 10, 'color' => '#64748B'],
                'fill' => '#F8FAFC',
                'alignment' => ['horizontal' => 'center', 'vertical' => 'center'],
            ],
            'mnb.currency' => [
                'format' => '$#,##0.00',
                'alignment' => ['horizontal' => 'right'],
            ],
            'mnb.percent' => [
                'format' => '0.00%',
                'alignment' => ['horizontal' => 'right'],
            ],
            'mnb.date' => [
                'format' => 'yyyy-mm-dd',
                'alignment' => ['horizontal' => 'left'],
            ],
            'mnb.datetime' => [
                'format' => 'yyyy-mm-dd hh:mm:ss',
                'alignment' => ['horizontal' => 'left'],
            ],
            'mnb.integer' => [
                'format' => 'integer',
                'alignment' => ['horizontal' => 'right'],
            ],
            'mnb.decimal' => [
                'format' => 'number',
                'alignment' => ['horizontal' => 'right'],
            ],
            'mnb.text' => [
                'format' => 'text',
                'alignment' => ['horizontal' => 'left'],
            ],
            'mnb.row.success' => [
                'fill' => '#ECFDF5',
                'font' => ['color' => '#064E3B'],
            ],
            'mnb.row.warning' => [
                'fill' => '#FFFBEB',
                'font' => ['color' => '#92400E'],
            ],
            'mnb.row.danger' => [
                'fill' => '#FEF2F2',
                'font' => ['color' => '#991B1B'],
            ],
            'mnb.row.muted' => [
                'fill' => '#F8FAFC',
                'font' => ['color' => '#64748B'],
            ],
        ];
    }
}
