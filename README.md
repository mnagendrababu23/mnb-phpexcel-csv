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
