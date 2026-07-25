<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

/**
 * Lightweight character-encoding detector and converter for CSV/text workflows.
 *
 * It intentionally avoids hard dependency on mbstring/iconv. When available,
 * mbstring/iconv are used for better conversion support; otherwise the class
 * falls back to safe UTF-8 / Windows-1252 / ISO-8859-1 handling.
 */
final class EncodingDetector
{
    /** @return list<string> */
    public static function defaultCandidates(): array
    {
        return [
            'UTF-8',
            'UTF-8-BOM',
            'UTF-16LE',
            'UTF-16BE',
            'UTF-32LE',
            'UTF-32BE',
            'Windows-1252',
            'ISO-8859-1',
        ];
    }

    /**
     * Detect text encoding from a binary sample.
     *
     * @param list<string>|null $candidates
     * @return array{encoding:string,confidence:float,source:string,bom:bool,notes:list<string>}
     */
    public static function detect(string $bytes, ?array $candidates = null): array
    {
        $notes = [];

        if ($bytes === '') {
            return [
                'encoding' => 'UTF-8',
                'confidence' => 0.25,
                'source' => 'empty',
                'bom' => false,
                'notes' => ['Empty sample; defaulted to UTF-8.'],
            ];
        }

        $bom = self::detectBom($bytes);
        if ($bom !== null) {
            return [
                'encoding' => $bom,
                'confidence' => 1.0,
                'source' => 'bom',
                'bom' => true,
                'notes' => ['Detected byte order mark.'],
            ];
        }

        $nulInfo = self::detectUtf16Or32ByNulPattern($bytes);
        if ($nulInfo !== null) {
            return $nulInfo;
        }

        if (self::isValidUtf8($bytes)) {
            return [
                'encoding' => 'UTF-8',
                'confidence' => self::looksAsciiOnly($bytes) ? 0.65 : 0.93,
                'source' => 'utf8-validation',
                'bom' => false,
                'notes' => self::looksAsciiOnly($bytes) ? ['ASCII-only sample is UTF-8 compatible.'] : ['Sample validates as UTF-8.'],
            ];
        }

        $candidateList = $candidates !== null && $candidates !== [] ? $candidates : self::defaultCandidates();
        $candidateList = array_values(array_unique(array_map([self::class, 'normalizeEncodingName'], $candidateList)));

        if (function_exists('mb_detect_encoding')) {
            $mbCandidates = array_values(array_filter($candidateList, static fn (string $name): bool => !str_contains($name, 'BOM')));
            $detected = @mb_detect_encoding($bytes, $mbCandidates, true);
            if (is_string($detected) && $detected !== '') {
                return [
                    'encoding' => self::normalizeEncodingName($detected),
                    'confidence' => 0.82,
                    'source' => 'mb_detect_encoding',
                    'bom' => false,
                    'notes' => ['Detected using mb_detect_encoding().'],
                ];
            }
            $notes[] = 'mb_detect_encoding() did not return a strict match.';
        }

        $controlBytes = preg_match_all('/[\x80-\x9F]/', $bytes) ?: 0;
        if ($controlBytes > 0) {
            return [
                'encoding' => 'Windows-1252',
                'confidence' => 0.70,
                'source' => 'heuristic',
                'bom' => false,
                'notes' => array_merge($notes, ['Bytes in 0x80-0x9F range suggest Windows-1252 over ISO-8859-1.']),
            ];
        }

        return [
            'encoding' => 'ISO-8859-1',
            'confidence' => 0.58,
            'source' => 'fallback',
            'bom' => false,
            'notes' => array_merge($notes, ['Sample is not valid UTF-8; defaulted to ISO-8859-1 fallback.']),
        ];
    }

    /**
     * Detect file encoding from a sample without loading the full file.
     *
     * @param list<string>|null $candidates
     * @return array{encoding:string,confidence:float,source:string,bom:bool,notes:list<string>,sample_bytes:int,path:string}
     */
    public static function detectFile(string $path, int $sampleBytes = 65536, ?array $candidates = null): array
    {
        if (!is_file($path)) {
            throw new MnbExcelException('File not found for encoding detection: ' . $path);
        }

        $handle = fopen($path, 'rb');
        if ($handle === false) {
            throw new MnbExcelException('Unable to open file for encoding detection: ' . $path);
        }

        $sample = fread($handle, max(1024, $sampleBytes));
        fclose($handle);

        if (!is_string($sample)) {
            $sample = '';
        }

        $result = self::detect($sample, $candidates);
        $result['sample_bytes'] = strlen($sample);
        $result['path'] = $path;

        return $result;
    }

    public static function convert(string $value, string $targetEncoding = 'UTF-8', string $sourceEncoding = 'auto'): string
    {
        $targetEncoding = self::normalizeEncodingName($targetEncoding);
        $sourceEncoding = strtolower($sourceEncoding) === 'auto'
            ? self::detect($value)['encoding']
            : self::normalizeEncodingName($sourceEncoding);

        $value = self::stripBom($value, $sourceEncoding);

        if ($sourceEncoding === $targetEncoding || ($sourceEncoding === 'UTF-8-BOM' && $targetEncoding === 'UTF-8')) {
            return $value;
        }

        if (function_exists('mb_convert_encoding')) {
            $converted = @mb_convert_encoding($value, $targetEncoding, $sourceEncoding === 'UTF-8-BOM' ? 'UTF-8' : $sourceEncoding);
            if (is_string($converted)) {
                return $converted;
            }
        }

        if (function_exists('iconv')) {
            $converted = @iconv($sourceEncoding === 'UTF-8-BOM' ? 'UTF-8' : $sourceEncoding, $targetEncoding . '//TRANSLIT//IGNORE', $value);
            if (is_string($converted)) {
                return $converted;
            }
        }

        if ($targetEncoding === 'UTF-8' && in_array($sourceEncoding, ['Windows-1252', 'ISO-8859-1'], true)) {
            return self::latinLikeToUtf8($value, $sourceEncoding);
        }

        return $value;
    }

    /**
     * Convert a text file to UTF-8 temporary file and return the temp path.
     * The caller should unlink the returned file when done.
     */
    public static function convertFileToUtf8(string $path, string $sourceEncoding = 'auto'): string
    {
        if (!is_file($path)) {
            throw new MnbExcelException('File not found for UTF-8 conversion: ' . $path);
        }

        if (strtolower($sourceEncoding) === 'auto') {
            $sourceEncoding = self::detectFile($path)['encoding'];
        }
        $sourceEncoding = self::normalizeEncodingName($sourceEncoding);

        $contents = file_get_contents($path);
        if (!is_string($contents)) {
            throw new MnbExcelException('Unable to read file for UTF-8 conversion: ' . $path);
        }

        $converted = self::convert($contents, 'UTF-8', $sourceEncoding);
        $tempFile = tempnam(sys_get_temp_dir(), 'mnb_csv_utf8_');
        if ($tempFile === false) {
            throw new MnbExcelException('Unable to create temp file for UTF-8 conversion.');
        }

        file_put_contents($tempFile, $converted);
        return $tempFile;
    }

    public static function normalizeEncodingName(string $encoding): string
    {
        $encoding = strtoupper(trim(str_replace('_', '-', $encoding)));
        return match ($encoding) {
            'AUTO' => 'auto',
            'UTF8' => 'UTF-8',
            'UTF-8-SIG', 'UTF-8-BOM' => 'UTF-8-BOM',
            'UTF16', 'UTF-16' => 'UTF-16LE',
            'UTF16LE', 'UTF-16LE' => 'UTF-16LE',
            'UTF16BE', 'UTF-16BE' => 'UTF-16BE',
            'UTF32', 'UTF-32' => 'UTF-32LE',
            'UTF32LE', 'UTF-32LE' => 'UTF-32LE',
            'UTF32BE', 'UTF-32BE' => 'UTF-32BE',
            'CP1252', 'WINDOWS1252', 'WINDOWS-1252' => 'Windows-1252',
            'LATIN1', 'LATIN-1', 'ISO8859-1', 'ISO-8859-1' => 'ISO-8859-1',
            default => $encoding,
        };
    }

    public static function isUtf16Or32(string $encoding): bool
    {
        $encoding = self::normalizeEncodingName($encoding);
        return in_array($encoding, ['UTF-16LE', 'UTF-16BE', 'UTF-32LE', 'UTF-32BE'], true);
    }

    private static function detectBom(string $bytes): ?string
    {
        return match (true) {
            str_starts_with($bytes, "\xEF\xBB\xBF") => 'UTF-8-BOM',
            str_starts_with($bytes, "\xFF\xFE\x00\x00") => 'UTF-32LE',
            str_starts_with($bytes, "\x00\x00\xFE\xFF") => 'UTF-32BE',
            str_starts_with($bytes, "\xFF\xFE") => 'UTF-16LE',
            str_starts_with($bytes, "\xFE\xFF") => 'UTF-16BE',
            default => null,
        };
    }

    /** @return array{encoding:string,confidence:float,source:string,bom:bool,notes:list<string>}|null */
    private static function detectUtf16Or32ByNulPattern(string $bytes): ?array
    {
        $length = strlen($bytes);
        if ($length < 8) {
            return null;
        }

        $sampleLength = min($length, 4096);
        $sample = substr($bytes, 0, $sampleLength);

        $evenNul = 0;
        $oddNul = 0;
        for ($i = 0; $i < $sampleLength; $i++) {
            if ($sample[$i] === "\0") {
                if (($i % 2) === 0) {
                    $evenNul++;
                } else {
                    $oddNul++;
                }
            }
        }

        $totalNul = $evenNul + $oddNul;
        if ($totalNul < 4) {
            return null;
        }

        if ($oddNul > ($evenNul * 4)) {
            return [
                'encoding' => 'UTF-16LE',
                'confidence' => 0.86,
                'source' => 'nul-pattern',
                'bom' => false,
                'notes' => ['NUL-byte pattern suggests UTF-16 little endian.'],
            ];
        }

        if ($evenNul > ($oddNul * 4)) {
            return [
                'encoding' => 'UTF-16BE',
                'confidence' => 0.86,
                'source' => 'nul-pattern',
                'bom' => false,
                'notes' => ['NUL-byte pattern suggests UTF-16 big endian.'],
            ];
        }

        return null;
    }

    private static function isValidUtf8(string $bytes): bool
    {
        return preg_match('//u', $bytes) === 1;
    }

    private static function looksAsciiOnly(string $bytes): bool
    {
        return preg_match('/[^\x00-\x7F]/', $bytes) !== 1;
    }

    private static function stripBom(string $value, string $encoding): string
    {
        return match (self::normalizeEncodingName($encoding)) {
            'UTF-8-BOM' => preg_replace('/^\xEF\xBB\xBF/', '', $value) ?? $value,
            'UTF-16LE' => str_starts_with($value, "\xFF\xFE") ? substr($value, 2) : $value,
            'UTF-16BE' => str_starts_with($value, "\xFE\xFF") ? substr($value, 2) : $value,
            'UTF-32LE' => str_starts_with($value, "\xFF\xFE\x00\x00") ? substr($value, 4) : $value,
            'UTF-32BE' => str_starts_with($value, "\x00\x00\xFE\xFF") ? substr($value, 4) : $value,
            default => $value,
        };
    }

    private static function latinLikeToUtf8(string $value, string $encoding): string
    {
        $map1252 = [
            "\x80" => "€", "\x82" => "‚", "\x83" => "ƒ", "\x84" => "„", "\x85" => "…", "\x86" => "†", "\x87" => "‡",
            "\x88" => "ˆ", "\x89" => "‰", "\x8A" => "Š", "\x8B" => "‹", "\x8C" => "Œ", "\x8E" => "Ž",
            "\x91" => "‘", "\x92" => "’", "\x93" => "“", "\x94" => "”", "\x95" => "•", "\x96" => "–", "\x97" => "—",
            "\x98" => "˜", "\x99" => "™", "\x9A" => "š", "\x9B" => "›", "\x9C" => "œ", "\x9E" => "ž", "\x9F" => "Ÿ",
        ];

        $chars = preg_split('//', $value, -1, PREG_SPLIT_NO_EMPTY);
        if ($chars === false) {
            return $value;
        }

        $out = '';
        foreach ($chars as $char) {
            $ord = ord($char);
            if ($ord < 128) {
                $out .= $char;
                continue;
            }

            if ($encoding === 'Windows-1252' && isset($map1252[$char])) {
                $out .= $map1252[$char];
                continue;
            }

            $out .= self::codePointToUtf8($ord);
        }

        return $out;
    }
    private static function codePointToUtf8(int $code): string
    {
        if ($code <= 0x7F) {
            return chr($code);
        }
        if ($code <= 0x7FF) {
            return chr(0xC0 | ($code >> 6)) . chr(0x80 | ($code & 0x3F));
        }
        if ($code <= 0xFFFF) {
            return chr(0xE0 | ($code >> 12)) . chr(0x80 | (($code >> 6) & 0x3F)) . chr(0x80 | ($code & 0x3F));
        }

        return '';
    }

}
