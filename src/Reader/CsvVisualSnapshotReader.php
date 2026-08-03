<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Snapshot\VisualSnapshot;
use Mnb\PHPExcel\Snapshot\VisualSnapshotReaderInterface;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\MnbExcelException;

/** CSV has values and dialect, but no embedded workbook style model. */
final class CsvVisualSnapshotReader implements VisualSnapshotReaderInterface
{
    public function __construct(private readonly CsvReader $reader = new CsvReader())
    {
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function visualSnapshot(string $path, array $options = []): array
    {
        $meta = (new CsvMetadataReader())->metaInfo($path, array_replace(['profile' => 'full'], $options));
        $format = is_array($meta['format_details'] ?? null) ? $meta['format_details'] : [];
        $dialect = [
            'delimiter' => (string) ($format['delimiter'] ?? ','),
            'enclosure' => (string) ($format['enclosure'] ?? '"'),
            'escape' => (string) ($format['escape_character'] ?? '\\'),
            'line_ending' => $this->lineEnding((string) ($format['line_ending'] ?? 'LF')),
            'encoding' => (string) ($format['encoding'] ?? 'UTF-8'),
            'bom' => (bool) ($format['bom'] ?? false),
        ];
        $rows = $this->reader->iterateSheet($path, 1, array_replace($options, [
            'delimiter' => $dialect['delimiter'],
            'enclosure' => $dialect['enclosure'],
            'escape' => $dialect['escape'],
            'ignore_blank_lines' => false,
        ]));
        $cells = [];
        $maxColumns = 1;
        $maxRow = 1;
        $maximumCells = max(1, (int) ($options['max_cells'] ?? 1_000_000));
        $seenCells = 0;
        foreach ($rows as $rowIndex => $row) {
            if (!is_array($row)) {
                continue;
            }
            $rowNumber = (int) $rowIndex + 1;
            $maxRow = max($maxRow, $rowNumber);
            $maxColumns = max($maxColumns, count($row));
            $seenCells += count($row);
            if ($seenCells > $maximumCells) {
                throw new \Mnb\PHPExcel\Support\MnbExcelException('Visual snapshot exceeds max_cells.');
            }
            foreach (array_values($row) as $columnIndex => $value) {
                $coordinate = Coordinate::columnIndexToName($columnIndex + 1) . $rowNumber;
                $cells[$coordinate] = [
                    'type' => $value === null ? 'blank' : 'text',
                    'value' => $value,
                ];
            }
        }
        $sheetName = (string) ($options['sheet_name'] ?? 'Sheet1');
        $snapshot = [
            'schema' => VisualSnapshot::SCHEMA,
            'schema_version' => VisualSnapshot::VERSION,
            'format' => str_contains(strtolower((string) ($meta['format_variant'] ?? 'csv')), 'tsv') ? 'tsv' : 'csv',
            'source' => [
                'file_name' => basename($path),
                'size_bytes' => filesize($path) ?: 0,
            ],
            'workbook' => [
                'active_sheet' => 1,
                'active_sheet_name' => $sheetName,
                'metadata' => [],
                'dialect' => $dialect,
            ],
            'styles' => [],
            'sheets' => [[
                'index' => 1,
                'name' => $sheetName,
                'state' => 'visible',
                'dimension' => 'A1:' . Coordinate::columnIndexToName($maxColumns) . $maxRow,
                'cells' => $cells,
                'layout' => [
                    'column_widths' => [],
                    'row_heights' => [],
                    'freeze_panes' => ['rows' => 0, 'columns' => 0, 'top_left_cell' => null],
                    'auto_filter_range' => null,
                ],
                'merged_cells' => [],
                'comments' => [],
                'hyperlinks' => [],
                'images' => [],
                'conditional_formats' => [],
                'data_validations' => [],
                'capabilities' => [
                    'values' => 'available',
                    'styles' => 'not_applicable',
                    'layout' => 'not_applicable',
                ],
            ]],
            'capabilities' => [
                'values' => 'available',
                'dialect' => 'available',
                'styles' => 'not_applicable',
                'layout' => 'not_applicable',
                'round_trip' => 'values_and_dialect_only',
            ],
            'warnings' => ['CSV does not store Excel styles, worksheet layout, formulas, comments, images, or workbook metadata.'],
        ];
        VisualSnapshot::assertValid($snapshot);
        return $snapshot;
    }

    private function lineEnding(string $name): string
    {
        return match (strtoupper($name)) {
            'CRLF' => "\r\n",
            'CR' => "\r",
            default => "\n",
        };
    }
}
