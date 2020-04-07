<?php

namespace FFILibXlsxWriter;

use FFILibXlsxWriter\Exceptions\FFILibXlsxWriterException;

/**
 * Class Worksheet
 * @see http://libxlsxwriter.github.io/chartheet_8h.html
 */
class Chartsheet extends Sheet
{
    /**
     * @return string
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a9cb69c057990a4139f0a19b53b04b805
     */
    protected function getType(): string
    {
        return 'chartsheet';
    }
}
