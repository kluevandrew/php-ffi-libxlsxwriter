<?php


namespace FFILibXlsxWriter;

use FFI\CData;
use FFILibXlsxWriter\Style\Format;

class Worksheet
{
    protected Workbook $workbook;
    protected CData $cWorksheet;

    /**
     * Worksheet constructor.
     * @param Workbook $workbook
     * @param string|null $name
     */
    public function __construct(Workbook $workbook, ?string $name)
    {
        $this->workbook = $workbook;
        $this->cWorksheet = FFILibXlsxWriter::ffi()->workbook_add_worksheet(
            $this->workbook->getCWorkbook(),
            $name
        );
    }

    /**
     * @return CData
     */
    public function getCWorksheet(): CData
    {
        return $this->cWorksheet;
    }

    /**
     * @param int $row
     * @param int $col
     * @param string $string
     * @param Format|null $format
     *
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ac208046e7a6d12cc87982422efa41b31
     */
    public function writeString(int $row, int $col, string $string, Format $format = null) {
        FFILibXlsxWriter::ffi()->worksheet_write_string(
            $this->cWorksheet,
            $row,
            $col,
            $string,
            $format ? $format->getCFormat() : null
        );
    }

    /**
     * @param int $row
     * @param int $col
     * @param float $number
     * @param Format|null $format
     *
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ad9fc47d3beaa2ab4759414e8580c2289
     */
    public function writeNumber(int $row, int $col, float $number, Format $format = null) {
        FFILibXlsxWriter::ffi()->worksheet_write_number(
            $this->cWorksheet,
            $row,
            $col,
            $number,
            $format ? $format->getCFormat() : null
        );
    }

    /**
     * @param int $row
     * @param int $col
     * @param string $formula
     * @param Format|null $format
     *
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ae57117f04c82bef29805ec3eabc219bb
     */
    public function writeFormula(int $row, int $col, string $formula, Format $format = null) {
        FFILibXlsxWriter::ffi()->worksheet_write_formula(
            $this->cWorksheet,
            $row,
            $col,
            $formula,
            $format ? $format->getCFormat() : null
        );
    }
}
