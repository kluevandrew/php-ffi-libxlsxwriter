<?php

namespace FFILibXlsxWriter\Style;

use FFI\CData;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

class Format
{
    private Workbook $workbook;
    private CData $cFormat;

    /**
     * Format constructor.
     * @param Workbook $workbook
     */
    public function __construct(Workbook $workbook)
    {
        $this->workbook = $workbook;
        $this->cFormat = FFILibXlsxWriter::ffi()->workbook_add_format($this->workbook->getCWorkbook());
    }

    /**
     * @return CData
     */
    public function getCFormat(): CData
    {
        return $this->cFormat;
    }

    /**
     * @return $this
     *
     * @see http://libxlsxwriter.github.io/format_8h.html#a85e1b0baf44b445b65894e48722aec14
     */
    public function setBold(): self
    {
        FFILibXlsxWriter::ffi()->format_set_bold($this->cFormat);

        return $this;
    }

    /**
     * @param string $numFormat
     * @return $this
     *
     * @see http://libxlsxwriter.github.io/format_8h.html#af77bbd0003344cb16d455c7fb709e16c
     */
    public function setNumFormat(string $numFormat): self
    {
        FFILibXlsxWriter::ffi()->format_set_num_format($this->cFormat, $numFormat);

        return $this;
    }

}
