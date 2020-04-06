<?php

namespace FFILibXlsxWriter;

use FFILibXlsxWriter\Contracts\Closable;
use FFILibXlsxWriter\Contracts\DoNotFreeDirectly;
use FFILibXlsxWriter\Exceptions\FFILibXlsxWriterException;
use FFILibXlsxWriter\Exceptions\LibWarning;
use FFILibXlsxWriter\Structs\DataValidation;
use FFILibXlsxWriter\Structs\ImageBuffer;
use FFILibXlsxWriter\Options\ImageOptions;
use FFILibXlsxWriter\Structs\IntegerList;
use FFILibXlsxWriter\Structs\Protection;
use FFILibXlsxWriter\Structs\RichString;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Options\CommentOptions;
use FFILibXlsxWriter\Structs\DateTime;
use FFILibXlsxWriter\Traits\ClosableStruct;

class Worksheet extends Struct implements DoNotFreeDirectly, Closable
{
    use ClosableStruct;

    protected Workbook $workbook;

    private ?IntegerList $horizontalBreaks = null;

    private ?IntegerList $verticalBreaks = null;

    /**
     * Worksheet constructor.
     * @param Workbook $workbook
     * @param string|null $name
     */
    public function __construct(Workbook $workbook, ?string $name)
    {
        $this->workbook = $workbook;
        $this->pointer = FFILibXlsxWriter::ffi()->workbook_add_worksheet(
            $this->workbook->getPointer(),
            $name
        );
        $this->struct = $this->pointer[0];
    }

    protected function free(): void
    {
        $this->pointer = null;
        $this->struct = null;

        parent::free();
    }

    /**
     * @param int[]|string $cell
     * @param float $number
     * @param Format|null $format
     * @return Worksheet
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ad9fc47d3beaa2ab4759414e8580c2289
     */
    public function writeNumber($cell, float $number, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        FFILibXlsxWriter::ffi()->worksheet_write_number(
            $this->getPointer(),
            $row,
            $col,
            $number,
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param string $string
     * @param Format|null $format
     * @return Worksheet
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ac208046e7a6d12cc87982422efa41b31
     */
    public function writeString($cell, string $string, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        FFILibXlsxWriter::ffi()->worksheet_write_string(
            $this->getPointer(),
            $row,
            $col,
            $string,
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param string $formula
     * @param Format|null $format
     * @param float|null $result
     * @return Worksheet
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ae57117f04c82bef29805ec3eabc219bb
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af474c01f7ec8ac1a01206c9c4a5867af
     */
    public function writeFormula($cell, string $formula, Format $format = null, float $result = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        if (null !== $result) {
            FFILibXlsxWriter::ffi()->worksheet_write_formula_num(
                $this->getPointer(),
                $row,
                $col,
                $formula,
                $format ? $format->getPointer() : null,
                $result
            );
        } else {
            FFILibXlsxWriter::ffi()->worksheet_write_formula(
                $this->getPointer(),
                $row,
                $col,
                $formula,
                $format ? $format->getPointer() : null
            );
        }

        return $this;
    }

    /**
     * @param int[]|string $range
     * @param string $formula
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a8331799b5821d70c442d637ca66dd6e7
     */
    public function writeArrayFormula($range, string $formula, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list(
            $firstRow,
            $firstCol,
            $lastRow,
            $lastCol,
        ) = FFILibXlsxWriter::range($range);

        FFILibXlsxWriter::ffi()->worksheet_write_array_formula(
            $this->getPointer(),
            $firstRow,
            $firstCol,
            $lastRow,
            $lastCol,
            $formula,
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param DateTime $dateTime
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#afb0b3211160927a3412be28c9364b1b5
     */
    public function writeDatetime($cell, DateTime $dateTime, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        $this->workbook->addStructure($dateTime);

        FFILibXlsxWriter::ffi()->worksheet_write_datetime(
            $this->getPointer(),
            $row,
            $col,
            $dateTime->getPointer(),
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param string $url
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9b2ac96ee23574a432f5703eedcaf9a1
     */
    public function writeUrl($cell, string $url, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        FFILibXlsxWriter::ffi()->worksheet_write_url(
            $this->getPointer(),
            $row,
            $col,
            $url,
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param bool $value
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af14fc7b20ba84e56510caf15c577cf8c
     */
    public function writeBoolean($cell, bool $value, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        FFILibXlsxWriter::ffi()->worksheet_write_boolean(
            $this->getPointer(),
            $row,
            $col,
            (int)$value,
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param mixed $richString
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a62bf44845ce9dcc505bf228999db5afa
     */
    public function writeRichString($cell, RichString $richString, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        FFILibXlsxWriter::ffi()->worksheet_write_rich_string(
            $this->getPointer(),
            $row,
            $col,
            $richString->getStruct(),
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param string $comment
     * @param CommentOptions|null $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#abd1728d105365c0113e15f40c6bb1b27
     */
    public function writeComment($cell, string $comment, CommentOptions $options = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        if (null !== $options) {
            $this->workbook->addStructure($options);

            FFILibXlsxWriter::ffi()->worksheet_write_comment_opt(
                $this->getPointer(),
                $row,
                $col,
                $comment,
                $options->getPointer()
            );
        } else {
            FFILibXlsxWriter::ffi()->worksheet_write_comment(
                $this->getPointer(),
                $row,
                $col,
                $comment
            );
        }


        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a8c81b7e398d686a1fe17e76ef4485f34
     */
    public function writeBlank($cell, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        FFILibXlsxWriter::ffi()->worksheet_write_blank(
            $this->getPointer(),
            $row,
            $col,
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $cols
     * @param int $width
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9656e4e05d3787eee6b3e4d8e82d9b7f
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a52c4c145f684c5b4dcd2ed304d1fe907
     * @todo options
     */
    public function setColumn($cols, int $width, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list(
            $firstCol,
            $lastCol,
        ) = FFILibXlsxWriter::cols($cols);

        FFILibXlsxWriter::ffi()->worksheet_set_column(
            $this->getPointer(),
            $firstCol,
            $lastCol,
            $width,
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int $row
     * @param int $height
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ab9b7fb95e1bd9b0da70befd0d37a9173
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#abbc2de45e0aa84341fb10a98778b3807
     * @todo options
     */
    public function setRow(int $row, int $height, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);

        FFILibXlsxWriter::ffi()->worksheet_set_row(
            $this->getPointer(),
            $row,
            $height,
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $range
     * @param string $string
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ad5a2a09ec65c0f286b756235c7327225
     */
    public function mergeRange($range, string $string, Format $format = null): self
    {
        $this->throwIfClosed(__METHOD__);

        list($firstRow,
            $firstCol,
            $lastRow,
            $lastCol) = FFILibXlsxWriter::range($range);

        FFILibXlsxWriter::ffi()->worksheet_merge_range(
            $this->getPointer(),
            $firstRow,
            $firstCol,
            $lastRow,
            $lastCol,
            $string,
            $format ? $format->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $range
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a4e2b1de34e96331000a996f512aecfcf
     */
    public function autofilter($range): self
    {
        $this->throwIfClosed(__METHOD__);

        FFILibXlsxWriter::ffi()->worksheet_autofilter(
            $this->getPointer(),
            ...FFILibXlsxWriter::range($range)
        );

        return $this;
    }

    /**
     * @param string $password
     * @param Protection|null $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a1b49e135d4debcdb25941f2f40f04cb0
     */
    public function protect(
        string $password = null,
        Protection $options = null
    ): self {
        $this->throwIfClosed(__METHOD__);

        FFILibXlsxWriter::ffi()->worksheet_protect(
            $this->getPointer(),
            $password,
            $options ? $options->getPointer() : null
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a7e52f1eecb20fe4f9e6223bcf195103b
     */
    public function freezePanes($cell): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        FFILibXlsxWriter::ffi()->worksheet_freeze_panes(
            $this->getPointer(),
            $row,
            $col
        );

        return $this;
    }

    /**
     * @param int[]|string $range
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a62368aa313184d72a9ca2b1cf5de9a8a
     */
    public function setSelection($range): self
    {
        $this->throwIfClosed(__METHOD__);

        FFILibXlsxWriter::ffi()->worksheet_set_selection(
            $this->getPointer(),
            ...FFILibXlsxWriter::range($range)
        );

        return $this;
    }

    /**
     * @param float $vertical
     * @param float $horizontal
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9f4a3845529bcc2922b89bdb450ded32
     */
    public function splitPanes(float $vertical, float $horizontal): self
    {
        $this->throwIfClosed(__METHOD__);

        FFILibXlsxWriter::ffi()->worksheet_split_panes(
            $this->getPointer(),
            $vertical,
            $horizontal
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param string $path
     * @param ImageOptions|null $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a4529d77bcefcf478b8209f46fe730f6f
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a0b05e75e2c2a5c3452374714cdb2b79b
     */
    public function insertImage($cell, string $path, ImageOptions $options = null): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        if (null !== $options) {
            $this->workbook->addStructure($options);

            FFILibXlsxWriter::ffi()->worksheet_insert_image_opt(
                $this->getPointer(),
                $row,
                $col,
                $path,
                $options->getPointer()
            );
        } else {
            FFILibXlsxWriter::ffi()->worksheet_insert_image(
                $this->getPointer(),
                $row,
                $col,
                $path
            );
        }

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param ImageBuffer $imageBuffer
     * @param int $imageSize
     * @param ImageOptions|null $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#aebd5cc71a42ab0e4a9ce45fe9a6f6908
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af4c868a5cf0eaab8740c1f966ba5561c
     */
    public function insertImageBuffer(
        $cell,
        ImageBuffer $imageBuffer,
        int $imageSize,
        ImageOptions $options = null
    ): self {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        $this->workbook->addStructure($imageBuffer);

        if (null !== $options) {
            $this->workbook->addStructure($options);

            FFILibXlsxWriter::ffi()->worksheet_insert_image_buffer_opt(
                $this->getPointer(),
                $row,
                $col,
                $imageBuffer->getStruct(),
                $imageSize,
                $options->getPointer()
            );
        } else {
            FFILibXlsxWriter::ffi()->worksheet_insert_image_buffer(
                $this->getPointer(),
                $row,
                $col,
                $imageBuffer->getStruct(),
                $imageSize
            );
        }

        return $this;
    }

    /**
     * @param int $color
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a1e84ef4dff791fc2278dfb029af94cb0
     */
    public function setTabColor(int $color): self
    {
        $this->throwIfClosed(__METHOD__);

        FFILibXlsxWriter::ffi()->worksheet_set_tab_color(
            $this->getPointer(),
            $color
        );

        return $this;
    }

    /**
     * @param int[]|string $cell
     * @param DataValidation $dataValidation
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a70c4b9cfb27b18258117545ac6ce6da0
     */
    public function dataValidationCell($cell, DataValidation $dataValidation): self
    {
        $this->throwIfClosed(__METHOD__);
        list ($row, $col) = FFILibXlsxWriter::cell($cell);

        $this->workbook->addStructure($dataValidation);

        $result = FFILibXlsxWriter::ffi()->worksheet_data_validation_cell(
            $this->getPointer(),
            $row,
            $col,
            $dataValidation->getPointer()
        );

        if ($result !== 0) {
            throw FFILibXlsxWriterException::byCode(
                FFILibXlsxWriter::getLog(),
                $result
            );
        }

        return $this;
    }

    /**
     * @param string $header
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a4070c24ed5107f33e94f30a1bf865ba9
     */
    public function setHeader(string $header): self
    {
        $this->throwIfClosed(__METHOD__);

        FFILibXlsxWriter::ffi()->worksheet_set_header(
            $this->getPointer(),
            $header
        );

        return $this;
    }

    /**
     * @param string $footer
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a57eb561cf3ab5e408a6612f0e379903a
     */
    public function setFooter(string $footer): self
    {
        $this->throwIfClosed(__METHOD__);

        FFILibXlsxWriter::ffi()->worksheet_set_footer(
            $this->getPointer(),
            $footer
        );

        return $this;
    }

    /**
     * @param array $breaks
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9601745a2e9e7b1e194b7f5283f197f0
     */
    public function setHorizontalBreaks(array $breaks): self
    {
        $this->throwIfClosed(__METHOD__);

        if (null !== $this->horizontalBreaks) {
            $this->horizontalBreaks->free();
            $this->horizontalBreaks = null;
        }

        if (!empty($breaks)) {
            $this->horizontalBreaks = new IntegerList('uint32_t', ...$breaks);
        }

        FFILibXlsxWriter::ffi()->worksheet_set_h_pagebreaks(
            $this->getPointer(),
            $this->horizontalBreaks !== null ? $this->horizontalBreaks->getStruct() : null
        );

        return $this;
    }

    /**
     * @param array $breaks
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a00621cf25b5a449a92bb9b3fc83327e6
     */
    public function setVerticalBreaks(array $breaks): self
    {
        $this->throwIfClosed(__METHOD__);

        if (null !== $this->horizontalBreaks) {
            $this->verticalBreaks->free();
            $this->verticalBreaks = null;
        }

        if (!empty($breaks)) {
            $this->verticalBreaks = new IntegerList('uint16_t', ...$breaks);
        }

        FFILibXlsxWriter::ffi()->worksheet_set_v_pagebreaks(
            $this->getPointer(),
            $this->verticalBreaks !== null ? $this->verticalBreaks->getStruct() : null
        );

        return $this;
    }
}
