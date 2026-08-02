<?php

declare(strict_types=1);

$root = dirname(__DIR__);
$workspace = dirname($root);
spl_autoload_register(static function (string $class) use ($root, $workspace): void {
    $prefix = 'Mnb\\PHPExcel\\';
    if (!str_starts_with($class, $prefix)) {
        return;
    }
    $relative = str_replace('\\', '/', substr($class, strlen($prefix))) . '.php';
    foreach ([$root . '/src/', $workspace . '/mnb-phpexcel-core/src/'] as $base) {
        if (is_file($base . $relative)) {
            require $base . $relative;
            return;
        }
    }
});

use Mnb\PHPExcel\Format\Csv;

$runtime = __DIR__ . '/runtime';
@mkdir($runtime, 0775, true);
$path = $runtime . '/metadata.csv';
file_put_contents($path, "Name,Age,Link\r\nAlice,30,=HYPERLINK(\"https://example.com\")\r\nBob,25,ok\r\nBad,row\r\n");

$quick = Csv::metaInfo($path, ['profile' => 'quick', 'metadata_sample_rows' => 2]);
assert($quick['format'] === 'csv');
assert($quick['statistics']['state'] === 'partial');
assert($quick['statistics']['row_count'] === null);
assert($quick['format_details']['delimiter'] === ',');

$full = Csv::metaInfo($path, ['profile' => 'full']);
assert($full['statistics']['row_count'] === 4);
assert($full['statistics']['maximum_columns'] === 3);
assert($full['statistics']['ragged_row_count'] === 1);
assert($full['security']['formula_like_cell_count'] === 1);
assert($full['statistics']['header_detected'] === true);
assert($full['document']['state'] === 'not_applicable');
assert(count($full['security']['items']) === 1);

$session = Csv::read($path);
$viaSession = $session->metaInfo(['profile' => 'standard']);
assert($viaSession['statistics']['row_count'] === 4);

echo json_encode([
    'status' => 'passed',
    'rows' => $full['statistics']['row_count'],
    'ragged' => $full['statistics']['ragged_row_count'],
    'formula_like' => $full['security']['formula_like_cell_count'],
], JSON_PRETTY_PRINT) . PHP_EOL;
