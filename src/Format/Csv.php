<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Format;

use Mnb\PHPExcel\Core\WorkbookFactory;
use Mnb\PHPExcel\Reader\CsvMetadataReader;
use Mnb\PHPExcel\Reader\CsvReader;
use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\ReadSession;
use Mnb\PHPExcel\Writer\CsvWriter;

final class Csv
{
    /** @param array<string,mixed> $options @return array<string,mixed> */
    public static function metaInfo(string $path, array $options = []): array
    {
        return (new CsvMetadataReader())->metaInfo($path, $options);
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
