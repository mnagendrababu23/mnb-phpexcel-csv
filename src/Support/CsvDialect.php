<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

/**
 * CSV dialect presets for Excel-friendly and locale-friendly CSV files.
 */
final class CsvDialect
{
    /** @return array{delimiter:string,enclosure:string,escape:string,bom:bool,line_ending:string} */
    public static function resolve(array $options = []): array
    {
        $preset = strtolower((string) ($options['dialect'] ?? 'excel'));
        if ($preset === 'auto') {
            $preset = 'excel';
        }

        $dialects = [
            'excel' => [
                'delimiter' => ',',
                'enclosure' => '"',
                'escape' => '',
                'bom' => true,
                'line_ending' => "\r\n",
            ],
            'excel_tab' => [
                'delimiter' => "\t",
                'enclosure' => '"',
                'escape' => '',
                'bom' => true,
                'line_ending' => "\r\n",
            ],
            'semicolon' => [
                'delimiter' => ';',
                'enclosure' => '"',
                'escape' => '',
                'bom' => true,
                'line_ending' => "\r\n",
            ],
            'unix' => [
                'delimiter' => ',',
                'enclosure' => '"',
                'escape' => '',
                'bom' => false,
                'line_ending' => "\n",
            ],
            'pipe' => [
                'delimiter' => '|',
                'enclosure' => '"',
                'escape' => '',
                'bom' => false,
                'line_ending' => "\n",
            ],
        ];

        if (!isset($dialects[$preset])) {
            throw new MnbExcelException('Unknown CSV dialect: ' . $preset);
        }

        $resolved = $dialects[$preset];
        foreach (['delimiter', 'enclosure', 'escape', 'line_ending'] as $key) {
            if (array_key_exists($key, $options)) {
                $resolved[$key] = (string) $options[$key];
            }
        }
        if (array_key_exists('bom', $options)) {
            $resolved['bom'] = (bool) $options['bom'];
        }

        if ($resolved['delimiter'] === '') {
            throw new MnbExcelException('CSV delimiter cannot be empty.');
        }
        if ($resolved['enclosure'] === '') {
            throw new MnbExcelException('CSV enclosure cannot be empty.');
        }

        return $resolved;
    }

    /**
     * Detect the most likely delimiter by comparing field-count consistency
     * across a small sample of non-empty records.
     *
     * @param list<string>|null $candidates
     */
    public static function detectDelimiter(
        string $path,
        ?array $candidates = null,
        int $sampleLines = 12,
        string $enclosure = '"',
        string $escape = ''
    ): string {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('CSV file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }

        $candidates = $candidates ?? [',', ';', "\t", '|'];
        $candidates = array_values(array_filter(array_map('strval', $candidates), static fn(string $value): bool => $value !== ''));
        if ($candidates === []) {
            throw new MnbExcelException('CSV delimiter candidates cannot be empty.');
        }

        $handle = fopen($path, 'rb');
        if ($handle === false) {
            throw MnbExcelException::withCode('Unable to open CSV file: ' . $path, ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }

        $lines = [];
        try {
            while (count($lines) < max(2, $sampleLines) && ($line = fgets($handle)) !== false) {
                $line = preg_replace('/^\xEF\xBB\xBF/', '', $line) ?? $line;
                if (trim($line) !== '') {
                    $lines[] = $line;
                }
            }
        } finally {
            fclose($handle);
        }

        if ($lines === []) {
            return ',';
        }

        $best = ',';
        $bestScore = -INF;
        foreach ($candidates as $candidate) {
            $counts = [];
            foreach ($lines as $line) {
                $counts[] = count(str_getcsv($line, $candidate, $enclosure, $escape));
            }

            $frequency = array_count_values($counts);
            arsort($frequency);
            $modeCount = (int) array_key_first($frequency);
            $consistentRows = (int) reset($frequency);
            $variancePenalty = count(array_unique($counts)) - 1;
            $score = ($modeCount > 1 ? 1000 : 0) + ($consistentRows * 100) + $modeCount - ($variancePenalty * 10);

            if ($score > $bestScore) {
                $bestScore = $score;
                $best = $candidate;
            }
        }

        return $best;
    }

    private function __construct()
    {
    }
}
