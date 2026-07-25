<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Support\CsvDialect;
use Mnb\PHPExcel\Support\EncodingDetector;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\ValueSanitizer;

final class CsvWriter
{
    /**
     * @param array<string,mixed> $options
     */
    public function write(WorksheetData $sheet, string $path, array $options = []): void
    {
        $dialect = CsvDialect::resolve($options);
        $policy = (string) ($options['injection_policy'] ?? $options['csv_injection_policy'] ?? 'escape');
        $encoding = EncodingDetector::normalizeEncodingName((string) ($options['encoding'] ?? 'UTF-8'));
        if ($encoding === 'auto') {
            $encoding = 'UTF-8';
        }

        AtomicFileWriter::writeViaTemp($path, function (string $tmp) use ($sheet, $dialect, $policy, $encoding, $path): void {
            $handle = @fopen($tmp, 'wb');
            if ($handle === false) {
                throw MnbExcelException::withCode(
                    'Unable to open CSV file for writing: ' . $path,
                    ErrorCode::FILE_OPEN_FAILED,
                    ['path' => $path]
                );
            }

            try {
                if ($dialect['bom']) {
                    $this->writeBytes($handle, self::bomForEncoding($encoding), $path);
                }

                foreach ($sheet->rows as $row) {
                    $values = array_map(static function (mixed $value) use ($policy): string {
                        if ($value instanceof CellValue) {
                            $value = $value->displayValue();
                        }

                        if ($value === null) {
                            $value = '';
                        } elseif (is_bool($value)) {
                            $value = $value ? '1' : '0';
                        } else {
                            $value = (string) $value;
                        }

                        return (string) ValueSanitizer::sanitizeFormulaLikeText($value, $policy);
                    }, $row);

                    $line = self::buildCsvLine($values, $dialect['delimiter'], $dialect['enclosure'], $dialect['escape'], $dialect['line_ending']);
                    if ($encoding !== 'UTF-8') {
                        $line = EncodingDetector::convert($line, $encoding, 'UTF-8');
                    }
                    $this->writeBytes($handle, $line, $path);
                }
            } finally {
                if (!@fclose($handle)) {
                    throw MnbExcelException::withCode(
                        'Unable to close CSV file after writing: ' . $path,
                        ErrorCode::CSV_WRITE_FAILED,
                        ['path' => $path]
                    );
                }
            }
        });
    }

    /** @param list<string> $fields */
    private static function buildCsvLine(array $fields, string $delimiter, string $enclosure, string $escape, string $lineEnding): string
    {
        $temp = fopen('php://temp', 'r+b');
        if ($temp === false) {
            throw new MnbExcelException('Unable to create temporary CSV row buffer.');
        }

        $written = fputcsv($temp, $fields, $delimiter, $enclosure, $escape, $lineEnding);
        if ($written === false) {
            fclose($temp);
            throw new MnbExcelException('Unable to write CSV row.');
        }

        rewind($temp);
        $line = stream_get_contents($temp);
        fclose($temp);

        if (!is_string($line)) {
            throw new MnbExcelException('Unable to read temporary CSV row buffer.');
        }

        return $line;
    }


    /** @param resource $handle */
    private function writeBytes($handle, string $bytes, string $path): void
    {
        if ($bytes === '') {
            return;
        }

        $length = strlen($bytes);
        $written = 0;
        while ($written < $length) {
            $chunk = @fwrite($handle, substr($bytes, $written));
            if ($chunk === false || $chunk === 0) {
                throw MnbExcelException::withCode(
                    'Unable to write CSV file: ' . $path,
                    ErrorCode::CSV_WRITE_FAILED,
                    ['path' => $path, 'expected_bytes' => $length, 'written_bytes' => $written]
                );
            }
            $written += $chunk;
        }
    }

    private static function bomForEncoding(string $encoding): string
    {
        return match (EncodingDetector::normalizeEncodingName($encoding)) {
            'UTF-8', 'UTF-8-BOM' => "\xEF\xBB\xBF",
            'UTF-16LE' => "\xFF\xFE",
            'UTF-16BE' => "\xFE\xFF",
            'UTF-32LE' => "\xFF\xFE\x00\x00",
            'UTF-32BE' => "\x00\x00\xFE\xFF",
            default => '',
        };
    }
}
