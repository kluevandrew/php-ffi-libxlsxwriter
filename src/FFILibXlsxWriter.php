<?php

namespace FFILibXlsxWriter;

use FFI;
use FFI\CData;

class FFILibXlsxWriter
{
    /**
     * @var FFI|mixed
     */
    protected static ?FFI $ffi = null;

    /**
     * @var resource
     */
    protected static $stderr = null;

    /**
     * @var string
     */
    private static ?string $stderrFile = null;

    /**
     * FFILibXlsxWriter constructor.
     * @param string|null $libraryPath
     */
    public static function init(string $libraryPath = null)
    {
        if (null !== self::$ffi) {
            return;
        }

        self::$ffi = FFI::cdef(
            self::getHeaders(),
            $libraryPath ?? self::getLibraryPath(),
        );

        self::overloadLogging();
    }

    protected static function getHeaders(): string
    {
        return file_get_contents(__DIR__ . '/headers/xlsxwriter.h');
    }

    protected static function getLibraryPath(): string
    {
        return '/usr/local/lib/libxlsxwriter.so';
    }

    /**
     * @return FFI|CData|mixed Mixed is add for suppers PHPStorm warnings
     */
    public static function ffi(): FFI
    {
        return self::$ffi;
    }

    /**
     * @param array|string $address
     * @return array
     */
    public static function cell($address): array
    {
        if (is_array($address)) {
            return $address;
        }

        $ffi = self::ffi();
        $firstRow = $ffi->lxw_name_to_row($address);
        $firstCol = $ffi->lxw_name_to_col($address);

        return [$firstRow, $firstCol];
    }

    /**
     * @param array"string $range
     * @return array
     */
    public static function range($range): array
    {
        if (is_array($range)) {
            return $range;
        }

        $ffi = self::ffi();
        $firstRow = $ffi->lxw_name_to_row($range);
        $firstCol = $ffi->lxw_name_to_col($range);
        $lastRow = $ffi->lxw_name_to_row_2($range);
        $lastCol = $ffi->lxw_name_to_col_2($range);

        return [$firstRow, $firstCol, $lastRow, $lastCol];
    }

    /**
     * @param array|string $cols
     * @return array
     */
    public static function cols($cols): array
    {
        if (is_array($cols)) {
            return $cols;
        }

        $ffi = self::ffi();
        $firstCol = $ffi->lxw_name_to_col($cols);
        $lastCol = $ffi->lxw_name_to_col_2($cols);

        return [$firstCol, $lastCol];
    }

    /**
     * @param int|string $row
     * @return int
     */
    public static function col($row): int
    {
        if (is_int($row)) {
            return $row;
        }

        return self::ffi()->lxw_name_to_col($row);
    }

    /**
     * @void
     */
    private static function overloadLogging()
    {
        self::$stderrFile = tempnam(sys_get_temp_dir(), uniqid(__CLASS__, true));
        self::$ffi->freopen(self::$stderrFile, 'w', self::$ffi->stderr);
        self::$stderr = fopen(self::$stderrFile, 'r');
    }

    public static function getLog(): string
    {
        $log = [];
        while (!feof(self::$stderr)) {
            $log[] = fgets(self::$stderr);
        }

        return implode(PHP_EOL, $log);
    }

    public function __destruct()
    {
        if (self::$stderr !== null) {
            fclose(self::$stderr);
        }
        if (file_exists(self::$stderrFile)) {
            unlink(self::$stderrFile);
        }
    }
}
