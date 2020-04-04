<?php

namespace FFILibXlsxWriter;

use FFI\CData;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Structs\CommentOptions;
use FFILibXlsxWriter\Structs\DateTime;
use FFILibXlsxWriter\Structs\RichStringTupleCollection;

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
     * @param int $firstRow
     * @param int $firstCol
     * @param int $lastRow
     * @param int $lastCol
     * @param string $formula
     * @param Format|null $format
     * @return $this
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a8331799b5821d70c442d637ca66dd6e7
     */
    public function writeArrayFormula(
        int $firstRow,
        int $firstCol,
        int $lastRow,
        int $lastCol,
        string $formula,
        Format $format = null
    ): self
    {
        FFILibXlsxWriter::ffi()->worksheet_write_array_formula(
            $this->cWorksheet,
            $firstRow,
            $firstCol,
            $lastRow,
            $lastCol,
            $formula,
            $format ? $format->getCFormat() : null
        );

        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     * @param DateTime $dateTime
     * @param Format|null $format
     * @return $this
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#afb0b3211160927a3412be28c9364b1b5
     */
    public function writeDatetime(int $row, int $col, DateTime $dateTime, Format $format = null): self
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
     * @param int $row
     * @param int $col
     * @param string $url
     * @param Format|null $format
     * @return $this
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9b2ac96ee23574a432f5703eedcaf9a1
     */
    public function writeUrl(int $row, int $col, string $url, Format $format = null): self
    {
        FFILibXlsxWriter::ffi()->worksheet_write_url(
            $this->cWorksheet,
            $row,
            $col,
            $url,
            $format ? $format->getCFormat() : null
        );

        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     * @param bool $value
     * @param Format|null $format
     * @return $this
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af14fc7b20ba84e56510caf15c577cf8c
     */
    public function writeBoolean(int $row, int $col, bool $value, Format $format = null): self
    {
        FFILibXlsxWriter::ffi()->worksheet_write_boolean(
            $this->cWorksheet,
            $row,
            $col,
            (int)$value,
            $format ? $format->getCFormat() : null
        );

        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     * @param string $formula
     * @param Format|null $format
     * @param float $result
     * @return $this
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af474c01f7ec8ac1a01206c9c4a5867af
     * @noinspection PhpOptionalBeforeRequiredParametersInspection
     */
    public function writeFormulaNum(int $row, int $col, string $formula, Format $format = null, float $result): self
    {
        FFILibXlsxWriter::ffi()->worksheet_write_formula_num(
            $this->cWorksheet,
            $row,
            $col,
            $formula,
            $format ? $format->getCFormat() : null,
            $result
        );

        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     * @param mixed $richString
     * @param Format|null $format
     * @return $this
     * @throws \Exception
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a62bf44845ce9dcc505bf228999db5afa
     */
    public function writeRichString(
        int $row,
        int $col,
        string $richString,
        Format $format = null
    ): self
    {
        throw new \Exception('Rich strings are not implemented yet');
    }

    /**
     * @param int $row
     * @param int $col
     * @param string $comment
     * @return $this
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#abd1728d105365c0113e15f40c6bb1b27
     */
    public function writeComment(int $row, int $col, string $comment): self
    {
        FFILibXlsxWriter::ffi()->worksheet_write_comment(
            $this->cWorksheet,
            $row,
            $col,
            $comment
        );

        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     * @param string $comment
     * @param CommentOptions $options
     * @return $this
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a158dadac385cd1007994e5478b7b2aa7
     */
    public function writeCommentOpt(int $row, int $col, string $comment, CommentOptions $options): self
    {
        $this->workbook->addStructure($options);

        FFILibXlsxWriter::ffi()->worksheet_write_comment_opt(
            $this->cWorksheet,
            $row,
            $col,
            $comment,
            $options->getPointer()
        );

        return $this;
    }

    /**
     * @param int $row
     * @param int $col
     * @param Format|null $format
     * @return $this
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

    /**
     * @param int $firstCol
     * @param int $lastCol
     * @param int $width
     * @param Format|null $format
     * @return $this
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9656e4e05d3787eee6b3e4d8e82d9b7f
     */
    public function setColumn(int $firstCol, int $lastCol, int $width, Format $format = null): self
    {
        FFILibXlsxWriter::ffi()->worksheet_set_column(
            $this->cWorksheet,
            $firstCol,
            $lastCol,
            $width,
            $format ? $format->getCFormat() : null
        );

        return $this;
    }
}
