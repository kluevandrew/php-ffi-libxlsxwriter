<?php


namespace FFILibXlsxWriter;

use FFI\CData;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Types\LXWDateTime;

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
     * @param float $number
     * @param Format|null $format
     * @return Worksheet
     *
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ad9fc47d3beaa2ab4759414e8580c2289
     */
    public function writeNumber(int $row, int $col, float $number, Format $format = null): self
    {
        FFILibXlsxWriter::ffi()->worksheet_write_number(
            $this->cWorksheet,
            $row,
            $col,
            $number,
            $format ? $format->getCFormat() : null
        );

        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     * @param string $string
     * @param Format|null $format
     * @return Worksheet
     *
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ac208046e7a6d12cc87982422efa41b31
     */
    public function writeString(int $row, int $col, string $string, Format $format = null): self
    {
        FFILibXlsxWriter::ffi()->worksheet_write_string(
            $this->cWorksheet,
            $row,
            $col,
            $string,
            $format ? $format->getCFormat() : null
        );

        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     * @param string $formula
     * @param Format|null $format
     * @return Worksheet
     *
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ae57117f04c82bef29805ec3eabc219bb
     */
    public function writeFormula(int $row, int $col, string $formula, Format $format = null): self
    {
        FFILibXlsxWriter::ffi()->worksheet_write_formula(
            $this->cWorksheet,
            $row,
            $col,
            $formula,
            $format ? $format->getCFormat() : null
        );

        return $this;
    }

    /**
     * @return $this
     * @throws \Exception
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a8331799b5821d70c442d637ca66dd6e7
     */
    public function writeArrayFormula(): self {
        throw new \Exception('unimplemented');
        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     * @param LXWDateTime $dateTime
     * @param Format|null $format
     * @return $this
     *
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#afb0b3211160927a3412be28c9364b1b5
     */
    public function writeDatetime(int $row, int $col, LXWDateTime $dateTime, Format $format = null): self
    {
        $this->workbook->addStructure($dateTime);

        FFILibXlsxWriter::ffi()->worksheet_write_datetime(
            $this->cWorksheet,
            $row,
            $col,
            $dateTime->getPointer(),
            $format ? $format->getCFormat() : null
        );

        return $this;
    }

    /**
     * @return $this
     * @throws \Exception
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9b2ac96ee23574a432f5703eedcaf9a1
     */
    public function writeUrl(): self {
        throw new \Exception('unimplemented');
        return $this;
    }

    /**
     * @return $this
     * @throws \Exception
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af14fc7b20ba84e56510caf15c577cf8c
     */
    public function writeBoolean(): self {
        throw new \Exception('unimplemented');
        return $this;
    }

    /**
     * @return $this
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af474c01f7ec8ac1a01206c9c4a5867af
     */
    public function writeFormulaNum(): self {
        throw new \Exception('unimplemented');
        return $this;
    }

    /**
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a62bf44845ce9dcc505bf228999db5afa
     */
    public function writeRichString(): self {
        throw new \Exception('unimplemented');
        return $this;
    }

    /**
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#abd1728d105365c0113e15f40c6bb1b27
     */
    public function writeComment(): self {
        throw new \Exception('unimplemented');
        return $this;
    }

    /**
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a158dadac385cd1007994e5478b7b2aa7
     */
    public function writeCommentOpt(): self {
        throw new \Exception('unimplemented');
        return $this;
    }


    /**
     * @param int $row
     * @param int $col
     * @param Format|null $format
     * @return $this
     *
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a8c81b7e398d686a1fe17e76ef4485f34
     */
    public function writeBlank(int $row, int $col, Format $format = null): self
    {
        FFILibXlsxWriter::ffi()->worksheet_write_blank(
            $this->cWorksheet,
            $row,
            $col,
            $format ? $format->getCFormat() : null
        );

        return $this;
    }
}
