<?php

namespace FFILibXlsxWriter;

use FFI;
use FFI\CData;

class FFILibXlsxWriter
{
    protected static FFI $ffi;

    /**
     * FFILibXlsxWriter constructor.
     * @param string|null $libraryPath
     */
    public static function init(string $libraryPath = null)
    {
        self::$ffi = FFI::cdef(
            self::getHeaders(),
            $libraryPath ?? self::getLibraryPath(),
        );
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

    public static function cell(string $address): array
    {
        $ffi = self::ffi();
        $firstRow = $ffi->lxw_name_to_row($address);
        $firstCol = $ffi->lxw_name_to_col($address);

        return [$firstRow, $firstCol];
    }

    public static function range(string $range): array
    {
        $ffi = self::ffi();
        $firstRow = $ffi->lxw_name_to_row($range);
        $firstCol = $ffi->lxw_name_to_col($range);
        $lastRow = $ffi->lxw_name_to_row_2($range);
        $lastCol = $ffi->lxw_name_to_col_2($range);

        return [$firstRow, $firstCol, $lastRow, $lastCol];
    }

    public static function cols(string $cols): array
    {
        $ffi = self::ffi();
        $firstCol = $ffi->lxw_name_to_col($cols);
        $lastCol = $ffi->lxw_name_to_col_2($cols);

        return [$firstCol, $lastCol];
    }
}
