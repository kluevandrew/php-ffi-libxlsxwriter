<?php

namespace FFILibXlsxWriter;

use FFI;

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
     * @return FFI|mixed Mixed is add for suppers PHPStorm warnings
     */
    public static function ffi(): FFI
    {
        return self::$ffi;
    }

}
