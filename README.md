# MNB PHPExcel CSV

Independent CSV reader/writer module. Requires only `mnb/mnb-phpexcel-core`.

```bash
composer require mnb/mnb-phpexcel-csv:^2.0
```

```php
use Mnb\PHPExcel\Format\Csv;

$rows = Csv::read('customers.csv')->withHeaderRow()->toArray();

Csv::write($rows, 'customers-export.csv', [
    'with_header' => true,
    'dialect' => 'excel',
]);
```

Supports streaming reads, dialect detection/configuration, encoding conversion, BOM handling, projection, and formula-injection policies.

## Unified metadata

```php
$quick = Csv::metaInfo('customers.csv', [
    'profile' => 'quick',
]);

$full = Csv::read('customers.csv')->metaInfo([
    'profile' => 'full',
    'max_items' => 1000,
]);
```

The CSV collector uses the shared metadata schema and reports:

- file size, filesystem dates, MIME type, and optional SHA-256;
- detected encoding, BOM, delimiter, enclosure, escape character, and line endings;
- CSV/TSV variant and mixed line-ending detection;
- complete or sampled row/cell statistics;
- minimum, maximum, and modal column counts;
- ragged and blank rows;
- header-detection confidence;
- formula-like cells that may create CSV-injection risk;
- a synthetic single worksheet for cross-format API consistency.

Workbook-only sections such as document properties, macros, charts, pivots, and protection return `not_applicable`. CSV has no standard embedded metadata property store, so `updateMetaInfo()` is intentionally not provided for CSV files.
