<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Support\CsvDialect;
use Mnb\PHPExcel\Support\EncodingDetector;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

final class CsvReader implements IterableReaderInterface, FormatAwareReaderInterface, MetadataReaderInterface
{
    public function format(): string
    {
        return 'csv';
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function metaInfo(string $path, array $options = []): array
    {
        return (new CsvMetadataReader())->metaInfo($path, $options);
    }

    /** @return list<list<mixed>> */
    public function readSheet(string $path, int|string $sheet = 1, array $options = []): array
    {
        return array_values(iterator_to_array($this->iterateSheet($path, $sheet, $options), true));
    }

    /** @return \Generator<int,list<mixed>> */
    public function iterateSheet(string $path, int|string $sheet = 1, array $options = []): iterable
    {
        if ($sheet !== 1 && $sheet !== '1' && $sheet !== 'Sheet1') {
            throw new MnbExcelException('CSV supports only one sheet.');
        }
        if (!is_file($path)) {
            throw MnbExcelException::withCode('CSV file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        $maxFileBytes = isset($options['max_file_bytes']) ? max(0, (int) $options['max_file_bytes']) : null;
        $fileSize = filesize($path);
        if ($maxFileBytes !== null && $fileSize !== false && $fileSize > $maxFileBytes) {
            throw MnbExcelException::withCode(
                'CSV file exceeds max_file_bytes. Size: ' . $fileSize . ', max_file_bytes: ' . $maxFileBytes,
                ErrorCode::FILE_READ_FAILED,
                ['path' => $path, 'size_bytes' => $fileSize, 'max_file_bytes' => $maxFileBytes]
            );
        }

        $targetEncoding = EncodingDetector::normalizeEncodingName((string) ($options['target_encoding'] ?? 'UTF-8'));
        $requestedEncoding = (string) ($options['encoding'] ?? $options['source_encoding'] ?? 'auto');
        $sourceEncoding = strtolower(trim($requestedEncoding)) === 'auto'
            ? EncodingDetector::detectFile($path, (int) ($options['encoding_sample_bytes'] ?? 65536), $options['encoding_candidates'] ?? null)['encoding']
            : EncodingDetector::normalizeEncodingName($requestedEncoding);

        $readPath = $path;
        $tempPath = null;

        // UTF-16/UTF-32 CSV has NUL bytes between ASCII characters, so fgetcsv()
        // should read a normalized UTF-8 temp file instead of the raw file.
        if ($targetEncoding === 'UTF-8' && (EncodingDetector::isUtf16Or32($sourceEncoding) || $sourceEncoding === 'UTF-8-BOM')) {
            $tempPath = EncodingDetector::convertFileToUtf8($path, $sourceEncoding);
            $readPath = $tempPath;
            $sourceEncoding = 'UTF-8';
        }

        $dialectOptions = $options;
        $autoDelimiter = strtolower((string) ($options['delimiter'] ?? '')) === 'auto'
            || strtolower((string) ($options['dialect'] ?? '')) === 'auto';
        if ($autoDelimiter) {
            unset($dialectOptions['delimiter']);
            $baseDialect = CsvDialect::resolve($dialectOptions);
            $dialectOptions['delimiter'] = CsvDialect::detectDelimiter(
                $readPath,
                isset($options['delimiter_candidates']) ? (array) $options['delimiter_candidates'] : null,
                (int) ($options['delimiter_sample_lines'] ?? 12),
                $baseDialect['enclosure'],
                $baseDialect['escape']
            );
        }
        $dialect = CsvDialect::resolve($dialectOptions);
        $this->assertSingleByteControl($dialect['delimiter'], 'delimiter');
        $this->assertSingleByteControl($dialect['enclosure'], 'enclosure');
        if ($dialect['escape'] !== '') {
            $this->assertSingleByteControl($dialect['escape'], 'escape');
        }

        $handle = fopen($readPath, 'rb');
        if ($handle === false) {
            if ($tempPath !== null && is_file($tempPath)) {
                @unlink($tempPath);
            }
            throw MnbExcelException::withCode('Unable to open CSV file: ' . $path, ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }

        $projection = ColumnProjection::fromOptions($options);
        $trimValues = (bool) ($options['trim_values'] ?? false);
        $emptyToNull = (bool) ($options['empty_strings_to_null'] ?? false);
        $ignoreBlankLines = (bool) ($options['ignore_blank_lines'] ?? false);
        $startRow = max(1, (int) ($options['start_row'] ?? 1));
        $endRow = isset($options['end_row']) ? max(1, (int) $options['end_row']) : null;
        $sourceSkipRows = max(0, (int) ($options['source_skip_rows'] ?? 0));
        $sourceLimitRows = isset($options['source_limit_rows']) ? max(0, (int) $options['source_limit_rows']) : null;
        $maxSourceRows = isset($options['max_source_rows']) ? max(0, (int) $options['max_source_rows']) : null;
        $maxColumns = isset($options['max_columns']) ? max(1, (int) $options['max_columns']) : null;
        $strictColumns = (bool) ($options['strict_column_count'] ?? false);
        $expectedColumns = isset($options['expected_columns']) ? max(0, (int) $options['expected_columns']) : null;
        $delivered = 0;
        $physicalRow = 0;
        $stoppedByRequestedBoundary = false;

        try {
            while (($row = fgetcsv($handle, 0, $dialect['delimiter'], $dialect['enclosure'], $dialect['escape'])) !== false) {
                $physicalRow++;

                if ($physicalRow < $startRow || $physicalRow <= $sourceSkipRows) {
                    continue;
                }
                if ($endRow !== null && $physicalRow > $endRow) {
                    $stoppedByRequestedBoundary = true;
                    break;
                }
                if ($sourceLimitRows !== null && $delivered >= $sourceLimitRows) {
                    $stoppedByRequestedBoundary = true;
                    break;
                }

                if ($ignoreBlankLines && $this->isBlankCsvRow($row)) {
                    continue;
                }

                $normalized = [];
                foreach ($row as $index => $value) {
                    if (is_string($value)) {
                        if ($index === 0) {
                            $value = preg_replace('/^\xEF\xBB\xBF/', '', $value) ?? $value;
                        }

                        if ($sourceEncoding !== $targetEncoding && $sourceEncoding !== 'UTF-8-BOM') {
                            $value = EncodingDetector::convert($value, $targetEncoding, $sourceEncoding);
                        }

                        if ($trimValues) {
                            $value = trim($value);
                        }
                    }

                    if ($emptyToNull && $value === '') {
                        $value = null;
                    }
                    $normalized[] = $value;
                }

                $columnCount = count($normalized);
                if ($maxColumns !== null && $columnCount > $maxColumns) {
                    throw MnbExcelException::withCode(
                        'CSV column limit exceeded on row ' . $physicalRow . '. Columns: ' . $columnCount . ', max_columns: ' . $maxColumns,
                        ErrorCode::FILE_READ_FAILED,
                        ['path' => $path, 'row' => $physicalRow, 'columns' => $columnCount, 'max_columns' => $maxColumns]
                    );
                }

                if ($expectedColumns === null && $strictColumns) {
                    $expectedColumns = $columnCount;
                }
                if ($strictColumns && $expectedColumns !== null && $columnCount !== $expectedColumns) {
                    throw MnbExcelException::withCode(
                        'CSV row ' . $physicalRow . ' has ' . $columnCount . ' columns; expected ' . $expectedColumns . '.',
                        ErrorCode::FILE_READ_FAILED,
                        ['path' => $path, 'row' => $physicalRow, 'columns' => $columnCount, 'expected_columns' => $expectedColumns]
                    );
                }

                $delivered++;
                if ($maxSourceRows !== null && $delivered > $maxSourceRows) {
                    throw MnbExcelException::withCode(
                        'CSV row limit exceeded. Rows: ' . $delivered . ', max_source_rows: ' . $maxSourceRows,
                        ErrorCode::FILE_READ_FAILED,
                        ['path' => $path, 'rows' => $delivered, 'max_source_rows' => $maxSourceRows]
                    );
                }
                /** @var list<mixed> $projected */
                $projected = array_values($projection->project($normalized));
                yield $physicalRow - 1 => $projected;
            }

            if (!$stoppedByRequestedBoundary && !feof($handle)) {
                throw MnbExcelException::withCode(
                    'CSV parsing stopped before end of file near row ' . ($physicalRow + 1) . '.',
                    ErrorCode::FILE_READ_FAILED,
                    ['path' => $path, 'row' => $physicalRow + 1]
                );
            }
        } finally {
            fclose($handle);
            if ($tempPath !== null && is_file($tempPath)) {
                @unlink($tempPath);
            }
        }
    }

    private function isBlankCsvRow(array $row): bool
    {
        foreach ($row as $value) {
            if ($value !== null && trim((string) $value) !== '') {
                return false;
            }
        }
        return true;
    }

    private function assertSingleByteControl(string $value, string $name): void
    {
        if (strlen($value) !== 1) {
            throw new MnbExcelException('CSV ' . $name . ' must be exactly one byte.');
        }
    }
}
