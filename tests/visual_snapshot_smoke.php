<?php

declare(strict_types=1);
require __DIR__ . '/../vendor/autoload.php';
use Mnb\PHPExcel\Format\Csv;
$dir = sys_get_temp_dir() . '/mnb-csv-visual-' . bin2hex(random_bytes(4));
mkdir($dir, 0775, true);
$source = $dir . '/source.tsv';
$out = $dir . '/copy.tsv';
try {
    file_put_contents($source, "Name\tValue\r\nAlice\t10\r\nBob\t20\r\n");
    $snapshot = Csv::visualSnapshot($source, ['delimiter' => "\t"]);
    assert($snapshot['format'] === 'tsv');
    assert($snapshot['capabilities']['styles'] === 'not_applicable');
    assert($snapshot['sheets'][0]['dimension'] === 'A1:B3');
    Csv::createFromVisualSnapshot($snapshot, $out);
    $copy = Csv::visualSnapshot($out, ['delimiter' => "\t"]);
    assert($copy['sheets'][0]['cells']['A2']['value'] === 'Alice');
    assert($copy['workbook']['dialect']['delimiter'] === "\t");
    echo "csv_visual_snapshot_smoke passed\n";
} finally {
    @unlink($source); @unlink($out); @rmdir($dir);
}
