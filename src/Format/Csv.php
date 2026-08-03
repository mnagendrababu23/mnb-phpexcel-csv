<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Format;

use Mnb\PHPExcel\Core\WorkbookFactory;
use Mnb\PHPExcel\Reader\CsvMetadataReader;
use Mnb\PHPExcel\Reader\CsvReader;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\ReadSession;
use Mnb\PHPExcel\Writer\CsvWriter;
use Mnb\PHPExcel\Reader\CsvVisualSnapshotReader;
use Mnb\PHPExcel\Snapshot\VisualSnapshot;

final class Csv
{
    /** @param array<string,mixed> $options @return array<string,mixed> */
    public static function metaInfo(string $path, array $options = []): array
    {
        return (new CsvMetadataReader())->metaInfo($path, $options);
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public static function visualSnapshot(string $path, array $options = []): array
    {
        return (new CsvVisualSnapshotReader())->visualSnapshot($path, $options);
    }

    /** @param array<string,mixed>|string $snapshot @param array<string,mixed> $options */
    public static function createFromVisualSnapshot(array|string $snapshot, string $path, array $options = []): void
    {
        if (is_string($snapshot)) {
            $snapshot = is_file($snapshot)
                ? VisualSnapshot::fromJson((string) file_get_contents($snapshot))
                : VisualSnapshot::fromJson($snapshot);
        }
        $workbook = VisualSnapshot::toWorkbookData($snapshot, $options);
        $dialect = is_array($snapshot['workbook']['dialect'] ?? null) ? $snapshot['workbook']['dialect'] : [];
        (new CsvWriter())->write($workbook->sheets[0], $path, array_replace($dialect, $options));
    }

    /** @param array<string,mixed>|ReaderOptions $options */
    public static function read(string $path, array|ReaderOptions $options = []): ReadSession
    {
        return new ReadSession($path, new CsvReader(), $options);
    }

    /** @param iterable<array<int|string,mixed>|mixed> $rows @param array<string,mixed> $options */
    public static function write(iterable $rows, string $path, array $options = []): void
    {
        $sheet = WorkbookFactory::worksheet(
            $rows,
            (string) ($options['sheet_name'] ?? 'Sheet1'),
            (bool) ($options['with_header'] ?? false)
        );
        (new CsvWriter())->write($sheet, $path, $options);
    }
}
