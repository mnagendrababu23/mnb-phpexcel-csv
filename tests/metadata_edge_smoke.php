<?php

declare(strict_types=1);

$root = dirname(__DIR__);
$workspace = dirname($root);
spl_autoload_register(static function (string $class) use ($root, $workspace): void {
    $prefix = 'Mnb\\PHPExcel\\';
    if (!str_starts_with($class, $prefix)) return;
    $relative = str_replace('\\', '/', substr($class, strlen($prefix))) . '.php';
    foreach ([$root . '/src/', $workspace . '/mnb-phpexcel-core/src/'] as $base) {
        if (is_file($base . $relative)) { require $base . $relative; return; }
    }
});

use Mnb\PHPExcel\Format\Csv;
use Mnb\PHPExcel\Support\MnbExcelException;

function check(bool $condition, string $message): void {
    if (!$condition) throw new RuntimeException($message);
}

$runtime = __DIR__ . '/runtime';
@mkdir($runtime, 0775, true);

$empty = $runtime . '/empty.csv';
file_put_contents($empty, '');
$emptyMeta = Csv::metaInfo($empty, ['profile' => 'full']);
check($emptyMeta['statistics']['row_count'] === 0, 'Empty CSV row count failed.');
check($emptyMeta['statistics']['maximum_columns'] === 0, 'Empty CSV column count failed.');

$tsv = $runtime . '/sample.tsv';
file_put_contents($tsv, "Name\tValue\nA\t1\nB\t2\n");
$tsvMeta = Csv::metaInfo($tsv, ['profile' => 'full', 'delimiter' => 'auto']);
check($tsvMeta['format_variant'] === 'tsv', 'TSV variant detection failed.');
check($tsvMeta['format_details']['delimiter'] === "\t", 'TSV delimiter detection failed.');
check($tsvMeta['statistics']['row_count'] === 3, 'TSV row count failed.');

$multiline = $runtime . '/multiline.csv';
file_put_contents($multiline, "id,notes\r\n1,\"line one\nline two\"\r\n2,ok\n");
$multiMeta = Csv::metaInfo($multiline, ['profile' => 'full']);
check($multiMeta['statistics']['row_count'] === 3, 'Quoted multiline record count failed.');
check($multiMeta['format_details']['mixed_line_endings'] === true, 'Mixed line ending detection failed.');

$utf8Text = "Name,City\r\nAlice,東京\r\n";
$utf16 = $runtime . '/utf16.csv';
$encoded = iconv('UTF-8', 'UTF-16LE', $utf8Text);
if (!is_string($encoded)) throw new RuntimeException('Unable to create UTF-16 fixture.');
file_put_contents($utf16, "\xFF\xFE" . $encoded);
$utf16Meta = Csv::metaInfo($utf16, ['profile' => 'full']);
check($utf16Meta['format_details']['encoding'] === 'UTF-16LE', 'UTF-16 encoding detection failed.');
check($utf16Meta['format_details']['bom'] === true, 'UTF-16 BOM detection failed.');
check($utf16Meta['statistics']['row_count'] === 2, 'UTF-16 row count failed.');
check($utf16Meta['statistics']['maximum_columns'] === 2, 'UTF-16 column count failed.');

$quick = Csv::metaInfo($tsv, ['profile' => 'quick', 'metadata_sample_rows' => 2]);
check($quick['statistics']['state'] === 'partial', 'Quick CSV state failed.');
check($quick['statistics']['row_count'] === null, 'Quick CSV total row count should be unknown.');
check($quick['statistics']['scanned_rows'] === 2, 'Quick CSV sample count failed.');

$exact = $runtime . '/exact-sample.csv';
file_put_contents($exact, "A,B\n1,2");
$exactMeta = Csv::metaInfo($exact, ['profile' => 'quick', 'metadata_sample_rows' => 2]);
check($exactMeta['statistics']['state'] === 'available', 'Complete quick scan was incorrectly marked partial.');
check($exactMeta['statistics']['row_count'] === 2, 'Complete quick scan lost the exact row count.');

$failed = false;
try {
    Csv::metaInfo($tsv, ['delimiter' => '::']);
} catch (MnbExcelException) {
    $failed = true;
}
check($failed, 'Invalid multi-byte delimiter was not rejected.');

$failed = false;
try {
    Csv::metaInfo($tsv, ['delimiter' => ',', 'enclosure' => ',']);
} catch (MnbExcelException) {
    $failed = true;
}
check($failed, 'Identical delimiter and enclosure were not rejected.');

echo json_encode([
    'status' => 'passed',
    'empty_rows' => $emptyMeta['statistics']['row_count'],
    'tsv_rows' => $tsvMeta['statistics']['row_count'],
    'utf16_rows' => $utf16Meta['statistics']['row_count'],
    'multiline_rows' => $multiMeta['statistics']['row_count'],
], JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE) . PHP_EOL;
