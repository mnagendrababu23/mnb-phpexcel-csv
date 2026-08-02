<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Metadata\MetadataCapabilities;
use Mnb\PHPExcel\Metadata\MetadataOptions;
use Mnb\PHPExcel\Metadata\MetadataProfile;
use Mnb\PHPExcel\Metadata\MetadataReport;
use Mnb\PHPExcel\Metadata\MetadataSectionState;
use Mnb\PHPExcel\Support\CsvDialect;
use Mnb\PHPExcel\Support\EncodingDetector;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\ValueSanitizer;

/** Native metadata collector for CSV/TSV and other delimiter-separated text files. */
final class CsvMetadataReader
{
    /**
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public function metaInfo(string $path, array $options = []): array
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('CSV file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        if (!is_readable($path)) {
            throw MnbExcelException::withCode('CSV file is not readable: ' . $path, ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }

        $metadataOptions = MetadataOptions::fromArray($options);
        $encoding = EncodingDetector::detectFile(
            $path,
            max(1024, (int) ($options['encoding_sample_bytes'] ?? 65536)),
            isset($options['encoding_candidates']) ? (array) $options['encoding_candidates'] : null
        );

        [$scanPath, $temporary] = $this->normalizedScanPath($path, (string) $encoding['encoding']);
        try {
            $dialect = $this->resolveDialect($scanPath, $options);
            $lineEndings = $this->detectLineEndings($scanPath, max(4096, (int) ($options['line_ending_sample_bytes'] ?? 262144)));
            $scan = $this->scan($scanPath, $dialect, $metadataOptions, $options);
        } finally {
            if ($temporary && is_file($scanPath)) {
                @unlink($scanPath);
            }
        }

        $extension = strtolower((string) pathinfo($path, PATHINFO_EXTENSION));
        $variant = $dialect['delimiter'] === "\t" || $extension === 'tsv' ? 'tsv' : 'csv';
        $mimeType = $variant === 'tsv' ? 'text/tab-separated-values' : 'text/csv';
        $report = new MetadataReport('csv', $variant, $mimeType, $metadataOptions->profile());

        $stat = @stat($path) ?: [];
        $report->setSection('file', MetadataSectionState::AVAILABLE, [
            'path' => $path,
            'resolved_path' => realpath($path) ?: $path,
            'name' => basename($path),
            'extension' => $extension,
            'size_bytes' => isset($stat['size']) ? (int) $stat['size'] : null,
            'filesystem_created_at' => isset($stat['ctime']) ? date(DATE_ATOM, (int) $stat['ctime']) : null,
            'filesystem_modified_at' => isset($stat['mtime']) ? date(DATE_ATOM, (int) $stat['mtime']) : null,
            'filesystem_accessed_at' => isset($stat['atime']) ? date(DATE_ATOM, (int) $stat['atime']) : null,
            'readable' => is_readable($path),
            'writable' => is_writable($path),
            'sha256' => $metadataOptions->includeHash() ? hash_file('sha256', $path) : null,
        ]);

        $formatWarnings = array_values(array_unique(array_merge(
            array_map('strval', $encoding['notes'] ?? []),
            $scan['warnings']
        )));
        $report->setSection('format_details', MetadataSectionState::AVAILABLE, [
            'encoding' => $encoding['encoding'],
            'encoding_confidence' => $encoding['confidence'],
            'encoding_detection_source' => $encoding['source'],
            'encoding_sample_bytes' => $encoding['sample_bytes'],
            'bom' => (bool) $encoding['bom'],
            'delimiter' => $dialect['delimiter'],
            'delimiter_display' => $this->displayControl($dialect['delimiter']),
            'enclosure' => $dialect['enclosure'],
            'escape_character' => $dialect['escape'],
            'line_ending' => $lineEndings['dominant'],
            'line_ending_counts' => $lineEndings['counts'],
            'mixed_line_endings' => $lineEndings['mixed'],
            'mime_type' => $mimeType,
            'warnings' => $formatWarnings,
        ]);

        $sheet = [
            'index' => 1,
            'name' => (string) ($options['sheet_name'] ?? 'Sheet1'),
            'state' => 'visible',
            'synthetic' => true,
            'row_count' => $scan['row_count'],
            'minimum_columns' => $scan['minimum_columns'],
            'maximum_columns' => $scan['maximum_columns'],
        ];
        $report->setSection('workbook', MetadataSectionState::AVAILABLE, [
            'name' => pathinfo($path, PATHINFO_FILENAME),
            'sheet_count' => 1,
            'active_sheet_index' => 1,
            'active_sheet_name' => $sheet['name'],
            'sheets' => [$sheet],
            'count' => 1,
            'items' => [$sheet],
            'synthetic' => true,
        ]);

        foreach ([
            'document', 'revision', 'application', 'custom_properties', 'macros',
            'named_objects', 'links', 'hidden_content', 'comments_notes',
            'tracked_changes', 'embedded_objects', 'calculation', 'print_settings',
            'validation', 'pivot_metadata', 'xml_metadata',
        ] as $section) {
            $report->setSection($section, MetadataSectionState::NOT_APPLICABLE);
        }

        $securityState = $scan['truncated'] ? MetadataSectionState::PARTIAL : MetadataSectionState::AVAILABLE;
        $report->setSection('security', $securityState, [
            'encrypted' => false,
            'password_protection' => false,
            'formula_injection_risk' => $scan['formula_like_cell_count'] > 0,
            'formula_like_cell_count' => $scan['formula_like_cell_count'],
            'items' => $scan['formula_like_cells'],
            'count' => $scan['formula_like_cell_count'],
            'truncated' => $scan['formula_items_truncated'],
            'warnings' => $scan['formula_like_cell_count'] > 0
                ? ['Formula-like text was found. Apply an explicit CSV injection policy before exporting or opening untrusted data.']
                : [],
        ]);

        $statisticsState = $scan['truncated'] ? MetadataSectionState::PARTIAL : MetadataSectionState::AVAILABLE;
        $report->setSection('statistics', $statisticsState, [
            'row_count' => $scan['row_count'],
            'scanned_rows' => $scan['scanned_rows'],
            'blank_row_count' => $scan['blank_row_count'],
            'non_blank_row_count' => $scan['non_blank_row_count'],
            'minimum_columns' => $scan['minimum_columns'],
            'maximum_columns' => $scan['maximum_columns'],
            'modal_column_count' => $scan['modal_column_count'],
            'consistent_column_count' => $scan['ragged_row_count'] === 0,
            'ragged_row_count' => $scan['ragged_row_count'],
            'ragged_rows' => $scan['ragged_rows'],
            'formula_like_cell_count' => $scan['formula_like_cell_count'],
            'empty_cell_count' => $scan['empty_cell_count'],
            'total_cell_count' => $scan['total_cell_count'],
            'header_detected' => $scan['header']['detected'],
            'header_confidence' => $scan['header']['confidence'],
            'header_reason' => $scan['header']['reason'],
            'truncated' => $scan['truncated'] || $scan['ragged_items_truncated'],
            'warnings' => $scan['truncated']
                ? ['Quick profile scanned a sample rather than the complete file.']
                : [],
        ]);

        $array = $report->toArray();
        $report->capabilities(MetadataCapabilities::fromReport($array));
        return $report->toArray();
    }

    /** @return array{0:string,1:bool} */
    private function normalizedScanPath(string $path, string $encoding): array
    {
        $normalized = EncodingDetector::normalizeEncodingName($encoding);
        if (EncodingDetector::isUtf16Or32($normalized) || $normalized === 'UTF-8-BOM') {
            return [EncodingDetector::convertFileToUtf8($path, $normalized), true];
        }
        return [$path, false];
    }

    /** @param array<string,mixed> $options @return array{delimiter:string,enclosure:string,escape:string,bom:bool,line_ending:string} */
    private function resolveDialect(string $path, array $options): array
    {
        $dialectOptions = $options;
        $requested = strtolower(trim((string) ($options['delimiter'] ?? 'auto')));
        $auto = $requested === '' || $requested === 'auto' || strtolower((string) ($options['dialect'] ?? '')) === 'auto';
        if ($auto) {
            unset($dialectOptions['delimiter']);
            $base = CsvDialect::resolve($dialectOptions);
            $dialectOptions['delimiter'] = CsvDialect::detectDelimiter(
                $path,
                isset($options['delimiter_candidates']) ? array_values(array_map('strval', (array) $options['delimiter_candidates'])) : null,
                max(2, (int) ($options['delimiter_sample_lines'] ?? 20)),
                $base['enclosure'],
                $base['escape']
            );
        }
        $dialect = CsvDialect::resolve($dialectOptions);
        if (strlen($dialect['delimiter']) !== 1) {
            throw new MnbExcelException('CSV metadata delimiter must be exactly one byte.');
        }
        if (strlen($dialect['enclosure']) !== 1) {
            throw new MnbExcelException('CSV metadata enclosure must be exactly one byte.');
        }
        if ($dialect['escape'] !== '' && strlen($dialect['escape']) !== 1) {
            throw new MnbExcelException('CSV metadata escape character must be empty or exactly one byte.');
        }
        if ($dialect['delimiter'] === $dialect['enclosure']) {
            throw new MnbExcelException('CSV delimiter and enclosure must be different characters.');
        }
        return $dialect;
    }

    /**
     * @param array{delimiter:string,enclosure:string,escape:string,bom:bool,line_ending:string} $dialect
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    private function scan(string $path, array $dialect, MetadataOptions $metadataOptions, array $options): array
    {
        $handle = @fopen($path, 'rb');
        if ($handle === false) {
            throw MnbExcelException::withCode('Unable to open CSV file: ' . $path, ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }

        $quickLimit = max(2, (int) ($options['metadata_sample_rows'] ?? 100));
        $scanLimit = $metadataOptions->profile() === MetadataProfile::QUICK ? $quickLimit : PHP_INT_MAX;
        $maxItems = $metadataOptions->maxItems();
        $rows = 0;
        $blankRows = 0;
        $emptyCells = 0;
        $totalCells = 0;
        $formulaLike = 0;
        $formulaItems = [];
        $formulaItemsTruncated = false;
        $columnCounts = [];
        $firstRows = [];
        $warnings = [];
        $hitLimit = false;

        try {
            while (($row = fgetcsv($handle, 0, $dialect['delimiter'], $dialect['enclosure'], $dialect['escape'])) !== false) {
                $rows++;
                if ($rows === 1 && isset($row[0]) && is_string($row[0])) {
                    $row[0] = preg_replace('/^\xEF\xBB\xBF/', '', $row[0]) ?? $row[0];
                }
                $columns = count($row);
                $columnCounts[$columns] = ($columnCounts[$columns] ?? 0) + 1;
                $isBlank = $this->isBlankRow($row);
                if ($isBlank) {
                    $blankRows++;
                }
                foreach ($row as $column => $value) {
                    $totalCells++;
                    $text = $value === null ? '' : (string) $value;
                    if ($text === '') {
                        $emptyCells++;
                    }
                    if (ValueSanitizer::isFormulaLikeText($text)) {
                        $formulaLike++;
                        if (count($formulaItems) < $maxItems && $metadataOptions->atLeast(MetadataProfile::FULL)) {
                            $formulaItems[] = [
                                'row' => $rows,
                                'column' => $column + 1,
                                'value_preview' => $this->preview($text),
                            ];
                        } else if ($metadataOptions->atLeast(MetadataProfile::FULL)) {
                            $formulaItemsTruncated = true;
                        }
                    }
                }
                if (count($firstRows) < 2 && !$isBlank) {
                    $firstRows[] = array_map(static fn(mixed $value): string => $value === null ? '' : (string) $value, $row);
                }
                if ($rows >= $scanLimit) {
                    $hitLimit = fgetc($handle) !== false;
                    break;
                }
            }
            if (!$hitLimit && !feof($handle)) {
                $warnings[] = 'CSV parsing stopped before the end of the file.';
            }
        } finally {
            fclose($handle);
        }

        if ($columnCounts === []) {
            $minimum = 0;
            $maximum = 0;
            $mode = 0;
            $ragged = 0;
        } else {
            $keys = array_map('intval', array_keys($columnCounts));
            $minimum = min($keys);
            $maximum = max($keys);
            arsort($columnCounts);
            $mode = (int) array_key_first($columnCounts);
            $ragged = $rows - (int) reset($columnCounts);
        }

        $raggedRows = [];
        $raggedItemsTruncated = false;
        if ($metadataOptions->atLeast(MetadataProfile::FULL) && $rows > 0 && $ragged > 0) {
            $handle = @fopen($path, 'rb');
            if ($handle !== false) {
                $rowNumber = 0;
                try {
                    while (($row = fgetcsv($handle, 0, $dialect['delimiter'], $dialect['enclosure'], $dialect['escape'])) !== false) {
                        $rowNumber++;
                        if (count($row) !== $mode) {
                            if (count($raggedRows) < $maxItems) {
                                $raggedRows[] = ['row' => $rowNumber, 'columns' => count($row), 'expected_columns' => $mode];
                            } else {
                                $raggedItemsTruncated = true;
                                break;
                            }
                        }
                        if ($hitLimit && $rowNumber >= $scanLimit) {
                            break;
                        }
                    }
                } finally {
                    fclose($handle);
                }
            }
        }

        return [
            'row_count' => $hitLimit ? null : $rows,
            'scanned_rows' => $rows,
            'blank_row_count' => $blankRows,
            'non_blank_row_count' => $rows - $blankRows,
            'minimum_columns' => $minimum,
            'maximum_columns' => $maximum,
            'modal_column_count' => $mode,
            'ragged_row_count' => $ragged,
            'ragged_rows' => $raggedRows,
            'ragged_items_truncated' => $raggedItemsTruncated,
            'empty_cell_count' => $emptyCells,
            'total_cell_count' => $totalCells,
            'formula_like_cell_count' => $formulaLike,
            'formula_like_cells' => $formulaItems,
            'formula_items_truncated' => $formulaItemsTruncated,
            'header' => $this->detectHeader($firstRows),
            'truncated' => $hitLimit,
            'warnings' => $warnings,
        ];
    }

    /** @param list<list<string>> $rows @return array{detected:bool,confidence:float,reason:string} */
    private function detectHeader(array $rows): array
    {
        if ($rows === []) {
            return ['detected' => false, 'confidence' => 0.0, 'reason' => 'The file has no non-blank rows.'];
        }
        $first = $rows[0];
        $nonEmpty = array_values(array_filter(array_map('trim', $first), static fn(string $value): bool => $value !== ''));
        if ($nonEmpty === []) {
            return ['detected' => false, 'confidence' => 0.1, 'reason' => 'The first non-blank row contains no labels.'];
        }
        $unique = count(array_unique(array_map('strtolower', $nonEmpty))) === count($nonEmpty);
        $firstTextRatio = $this->textRatio($first);
        $secondTextRatio = isset($rows[1]) ? $this->textRatio($rows[1]) : $firstTextRatio;
        $score = 0.25 + ($unique ? 0.25 : 0.0) + ($firstTextRatio * 0.30) + (max(0.0, $firstTextRatio - $secondTextRatio) * 0.40);
        $score = min(1.0, $score);
        return [
            'detected' => $score >= 0.60,
            'confidence' => round($score, 3),
            'reason' => $score >= 0.60
                ? 'The first non-blank row contains unique text labels and differs from the following data row.'
                : 'The first rows do not provide strong evidence of a header.',
        ];
    }

    /** @param list<string> $row */
    private function textRatio(array $row): float
    {
        $nonEmpty = 0;
        $text = 0;
        foreach ($row as $value) {
            $value = trim($value);
            if ($value === '') {
                continue;
            }
            $nonEmpty++;
            if (!$this->looksLikeDataScalar($value)) {
                $text++;
            }
        }
        return $nonEmpty === 0 ? 0.0 : $text / $nonEmpty;
    }

    private function looksLikeDataScalar(string $value): bool
    {
        if (is_numeric(str_replace([',', ' '], '', $value))) {
            return true;
        }
        if (in_array(strtolower($value), ['true', 'false', 'yes', 'no', 'null'], true)) {
            return true;
        }
        return preg_match('/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}(?:[ T].*)?$/', $value) === 1;
    }

    /** @param list<mixed> $row */
    private function isBlankRow(array $row): bool
    {
        foreach ($row as $value) {
            if ($value !== null && trim((string) $value) !== '') {
                return false;
            }
        }
        return true;
    }

    /** @return array{dominant:string,counts:array{CRLF:int,LF:int,CR:int},mixed:bool} */
    private function detectLineEndings(string $path, int $sampleBytes): array
    {
        $handle = @fopen($path, 'rb');
        if ($handle === false) {
            return ['dominant' => 'unknown', 'counts' => ['CRLF' => 0, 'LF' => 0, 'CR' => 0], 'mixed' => false];
        }
        $sample = fread($handle, $sampleBytes);
        fclose($handle);
        if (!is_string($sample)) {
            $sample = '';
        }
        $crlf = substr_count($sample, "\r\n");
        $withoutCrLf = str_replace("\r\n", '', $sample);
        $lf = substr_count($withoutCrLf, "\n");
        $cr = substr_count($withoutCrLf, "\r");
        $counts = ['CRLF' => $crlf, 'LF' => $lf, 'CR' => $cr];
        arsort($counts);
        $dominant = (int) reset($counts) > 0 ? (string) array_key_first($counts) : 'none';
        $used = count(array_filter($counts, static fn(int $count): bool => $count > 0));
        return ['dominant' => $dominant, 'counts' => ['CRLF' => $crlf, 'LF' => $lf, 'CR' => $cr], 'mixed' => $used > 1];
    }

    private function displayControl(string $value): string
    {
        return match ($value) {
            "\t" => '\\t',
            "\r" => '\\r',
            "\n" => '\\n',
            default => $value,
        };
    }

    private function preview(string $value): string
    {
        $value = preg_replace('/[\r\n\t]+/', ' ', $value) ?? $value;
        return ValueSanitizer::substring($value, 0, 120);
    }
}
