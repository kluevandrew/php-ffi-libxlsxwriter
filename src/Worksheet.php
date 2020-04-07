<?php

namespace FFILibXlsxWriter;

use FFILibXlsxWriter\Exceptions\FFILibXlsxWriterException;
use FFILibXlsxWriter\Exceptions\UnimplementedException;
use FFILibXlsxWriter\Options\HeaderFooterOptions;
use FFILibXlsxWriter\Options\RowColOptions;
use FFILibXlsxWriter\Structs\DataValidation;
use FFILibXlsxWriter\Structs\ImageBuffer;
use FFILibXlsxWriter\Options\ImageOptions;
use FFILibXlsxWriter\Structs\IntegerList;
use FFILibXlsxWriter\Structs\Protection;
use FFILibXlsxWriter\Structs\RichString;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Options\CommentOptions;
use FFILibXlsxWriter\Structs\DateTime;
use FFILibXlsxWriter\Enums;

/**
 * Class Worksheet
 * @see http://libxlsxwriter.github.io/worksheet_8h.html
 */
class Worksheet extends Sheet
{
    protected ?IntegerList $horizontalBreaks = null;

    protected ?IntegerList $verticalBreaks = null;

    /**
     * @return string
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a81d456b4f65a464e78e4a0030ecc3c2e
     */
    protected function getType(): string
    {
        return 'worksheet';
    }

    /**
     * @param int[]|string $cell
     * @param float $number
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ad9fc47d3beaa2ab4759414e8580c2289
     */
    public function writeNumber($cell, float $number, Format $format = null): self
    {
        return $this->proxy(
            __METHOD__,
            'write_number',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $number,
                $format ? $format->getPointer() : null
            ]
        );
    }

    /**
     * @param int[]|string $cell
     * @param string $string
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ac208046e7a6d12cc87982422efa41b31
     */
    public function writeString($cell, string $string, Format $format = null): self
    {
        return $this->proxy(
            __METHOD__,
            'write_string',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $string,
                $format ? $format->getPointer() : null
            ]
        );
    }

    /**
     * @param int[]|string $cell
     * @param string $formula
     * @param Format|null $format
     * @param float|null $result
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ae57117f04c82bef29805ec3eabc219bb
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af474c01f7ec8ac1a01206c9c4a5867af
     */
    public function writeFormula($cell, string $formula, Format $format = null, float $result = null): self
    {
        $arguments = [
            ...FFILibXlsxWriter::cell($cell),
            $formula,
            $format ? $format->getPointer() : null
        ];
        $cMethod = 'write_formula';

        if (null !== $result) {
            $arguments[] = $result;
            $cMethod .= '_num';
        }

        return $this->proxy(
            __METHOD__,
            $cMethod,
            true,
            $arguments
        );
    }

    /**
     * @param int[]|string $range
     * @param string $formula
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a8331799b5821d70c442d637ca66dd6e7
     */
    public function writeArrayFormula($range, string $formula, Format $format = null): self
    {
        return $this->proxy(
            __METHOD__,
            'write_array_formula',
            true,
            [
                ...FFILibXlsxWriter::range($range),
                $formula,
                $format ? $format->getPointer() : null
            ]
        );
    }

    /**
     * @param int[]|string $cell
     * @param DateTime $dateTime
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#afb0b3211160927a3412be28c9364b1b5
     */
    public function writeDatetime($cell, DateTime $dateTime, Format $format = null): self
    {
        return $this->proxy(
            __METHOD__,
            'write_datetime',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $dateTime->getPointer(),
                $format ? $format->getPointer() : null
            ]
        );
    }

    /**
     * @param int[]|string $cell
     * @param string $url
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9b2ac96ee23574a432f5703eedcaf9a1
     */
    public function writeUrl($cell, string $url, Format $format = null): self
    {
        return $this->proxy(
            __METHOD__,
            'write_url',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $url,
                $format ? $format->getPointer() : null
            ]
        );
    }

    /**
     * @param int[]|string $cell
     * @param bool $value
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af14fc7b20ba84e56510caf15c577cf8c
     */
    public function writeBoolean($cell, bool $value, Format $format = null): self
    {
        return $this->proxy(
            __METHOD__,
            'write_boolean',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $value,
                $format ? $format->getPointer() : null
            ]
        );
    }

    /**
     * @param int[]|string $cell
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a8c81b7e398d686a1fe17e76ef4485f34
     */
    public function writeBlank($cell, Format $format = null): self
    {
        return $this->proxy(
            __METHOD__,
            'write_blank',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $format ? $format->getPointer() : null
            ]
        );
    }

    /**
     * @param int[]|string $cell
     * @param mixed $richString
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a62bf44845ce9dcc505bf228999db5afa
     */
    public function writeRichString($cell, RichString $richString, Format $format = null): self
    {
        return $this->proxy(
            __METHOD__,
            'write_rich_string',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $richString->getStruct(),
                $format ? $format->getPointer() : null
            ]
        );
    }

    /**
     * @param int[]|string $cell
     * @param string $comment
     * @param CommentOptions|null $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#abd1728d105365c0113e15f40c6bb1b27
     */
    public function writeComment($cell, string $comment, CommentOptions $options = null): self
    {
        return $this->proxyOpt(
            __METHOD__,
            'write_comment',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $comment,
            ],
            $options
        );
    }

    /**
     * @param int $row
     * @param int $height
     * @param Format|null $format
     * @param RowColOptions $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ab9b7fb95e1bd9b0da70befd0d37a9173
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#abbc2de45e0aa84341fb10a98778b3807
     */
    public function setRow(int $row, int $height, Format $format = null, RowColOptions $options = null): self
    {
        return $this->proxyOpt(
            __METHOD__,
            'set_row',
            true,
            [
                $row,
                $height,
                $format ? $format->getPointer() : null
            ],
            $options
        );
    }

    /**
     * @param int[]|string $cols
     * @param int $width
     * @param Format|null $format
     * @param RowColOptions|null $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9656e4e05d3787eee6b3e4d8e82d9b7f
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a52c4c145f684c5b4dcd2ed304d1fe907
     */
    public function setColumn($cols, int $width, Format $format = null, RowColOptions $options = null): self
    {
        return $this->proxyOpt(
            __METHOD__,
            'set_column',
            true,
            [
                ...FFILibXlsxWriter::cols($cols),
                $width,
                $format ? $format->getPointer() : null
            ],
            $options
        );
    }

    /**
     * @param int[]|string $cell
     * @param string $path
     * @param ImageOptions|null $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a4529d77bcefcf478b8209f46fe730f6f
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a0b05e75e2c2a5c3452374714cdb2b79b
     */
    public function insertImage($cell, string $path, ImageOptions $options = null): self
    {
        return $this->proxyOpt(
            __METHOD__,
            'insert_image',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $path,
            ],
            $options
        );
    }

    /**
     * @param int[]|string $cell
     * @param ImageBuffer $imageBuffer
     * @param int $size
     * @param ImageOptions|null $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#aebd5cc71a42ab0e4a9ce45fe9a6f6908
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af4c868a5cf0eaab8740c1f966ba5561c
     */
    public function insertImageBuffer($cell, ImageBuffer $imageBuffer, int $size, ImageOptions $options = null): self
    {
        return $this->proxyOpt(
            __METHOD__,
            'insert_image_buffer',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $imageBuffer->getStruct(),
                $size
            ],
            $options
        );
    }

    /**
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ac2067faaacb8bfa6550b019e915900a2
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a4e8ea2614bc214fa0cc8992733c51f5f
     */
    public function insertChart(): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param int[]|string $range
     * @param string $string
     * @param Format|null $format
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ad5a2a09ec65c0f286b756235c7327225
     */
    public function mergeRange($range, string $string, Format $format = null): self
    {
        return $this->proxy(
            __METHOD__,
            'merge_range',
            true,
            [
                ...FFILibXlsxWriter::range($range),
                $string,
                $format ? $format->getPointer() : null
            ]
        );
    }

    /**
     * @param int[]|string $range
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a4e2b1de34e96331000a996f512aecfcf
     */
    public function autofilter($range): self
    {
        return $this->proxy(
            __METHOD__,
            'autofilter',
            true,
            FFILibXlsxWriter::range($range)
        );
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
        return $this->proxy(
            __METHOD__,
            'data_validation_cell',
            true,
            [
                ...FFILibXlsxWriter::cell($cell),
                $dataValidation->getPointer()
            ]
        );
    }

    /**
     * @param int[]|string $range
     * @param DataValidation $dataValidation
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a14f482010dd0f817e17001d01644a57f
     */
    public function dataValidationRange($range, DataValidation $dataValidation): self
    {
        return $this->proxy(
            __METHOD__,
            'data_validation_range',
            true,
            [
                ...FFILibXlsxWriter::range($range),
                $dataValidation->getPointer()
            ]
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a76ec76f91328c512d3d86a35642f0a08
     */
    public function activate(): self
    {
        return $this->proxy(
            __METHOD__,
            'activate'
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a78bfe3a786149deca69248db57da182f
     */
    public function select(): self
    {
        return $this->proxy(
            __METHOD__,
            'select'
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a180fcbcf012b5c165a01490d1128dc11
     */
    public function hide(): self
    {
        return $this->proxy(
            __METHOD__,
            'hide'
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a463c8d3247f56e0ec191b148355f42f0
     */
    public function setFirstSheet(): self
    {
        return $this->proxy(
            __METHOD__,
            'set_first_sheet'
        );
    }

    /**
     * @param int[]|string $cell
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a7e52f1eecb20fe4f9e6223bcf195103b
     */
    public function freezePanes($cell): self
    {
        return $this->proxy(
            __METHOD__,
            'freeze_panes',
            false,
            FFILibXlsxWriter::cell($cell)
        );
    }

    /**
     * @param float $vertical
     * @param float $horizontal
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9f4a3845529bcc2922b89bdb450ded32
     */
    public function splitPanes(float $vertical, float $horizontal): self
    {
        return $this->proxy(
            __METHOD__,
            'split_panes',
            false,
            [
                $vertical,
                $horizontal
            ]
        );
    }

    /**
     * @param int[]|string $range
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a62368aa313184d72a9ca2b1cf5de9a8a
     */
    public function setSelection($range): self
    {
        return $this->proxy(
            __METHOD__,
            'set_selection',
            false,
            FFILibXlsxWriter::range($range)
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ad24e2a01fccb881e4cc478a92aa095e6
     */
    public function setLandscape(): self
    {
        return $this->proxy(
            __METHOD__,
            'set_landscape'
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af498d1caa032633240af6d26ef80515a
     */
    public function setPortrait(): self
    {
        return $this->proxy(
            __METHOD__,
            'set_portrait'
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a5d442a52a3549d5d2324f13787b66f47
     */
    public function setPageView(): self
    {
        return $this->proxy(
            __METHOD__,
            'set_page_view'
        );
    }

    /**
     * @param int $paperType PaperType::*
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9f8af12017797b10c5ee68e04149032f
     */
    public function setPaper(int $paperType): self
    {
        return $this->proxy(
            __METHOD__,
            'set_paper',
            false,
            [$paperType]
        );
    }

    /**
     * @param float $left
     * @param float $right
     * @param float $top
     * @param float $bottom
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ab942bbc0493eaa5cf5696da30b6f872d
     */
    public function setMargins(float $left, float $right, float $top, float $bottom): self
    {
        return $this->proxy(
            __METHOD__,
            'set_margins',
            false,
            [
                $left,
                $right,
                $top,
                $bottom,
            ]
        );
    }

    /**
     * @param string $header
     * @param HeaderFooterOptions $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a4070c24ed5107f33e94f30a1bf865ba9
     */
    public function setHeader(string $header, HeaderFooterOptions $options = null): self
    {
        return $this->proxyOpt(
            __METHOD__,
            'set_header',
            true,
            [$header],
            $options
        );
    }

    /**
     * @param string $footer
     * @param HeaderFooterOptions $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a57eb561cf3ab5e408a6612f0e379903a
     */
    public function setFooter(string $footer, HeaderFooterOptions $options = null): self
    {
        return $this->proxyOpt(
            __METHOD__,
            'set_footer',
            true,
            [$footer],
            $options
        );
    }

    /**
     * @param array $breaks
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
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

        return $this->proxy(
            __METHOD__,
            'set_h_pagebreaks',
            true,
            [
                $this->horizontalBreaks !== null ? $this->horizontalBreaks->getStruct() : null
            ]
        );
    }

    /**
     * @param array $breaks
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a00621cf25b5a449a92bb9b3fc83327e6
     */
    public function setVerticalBreaks(array $breaks): self
    {
        $this->throwIfClosed(__METHOD__);

        if (null !== $this->verticalBreaks) {
            $this->verticalBreaks->free();
            $this->verticalBreaks = null;
        }

        if (!empty($breaks)) {
            $this->verticalBreaks = new IntegerList('uint16_t', ...$breaks);
        }

        return $this->proxy(
            __METHOD__,
            'set_v_pagebreaks',
            true,
            [
                $this->verticalBreaks !== null ? $this->verticalBreaks->getStruct() : null
            ]
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a40d32d247593e01c00b5c5c606fb4bb3
     */
    public function printAcross(): self
    {
        return $this->proxy(__METHOD__, 'print_across');
    }

    /**
     * @param int $scale
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a38e16dc0a2e75d7ec99dca21fbeec89b
     */
    public function setZoom(int $scale): self
    {
        return $this->proxy(
            __METHOD__,
            'set_zoom',
            false,
            [$scale]
        );
    }

    /**
     * @param int $option Gridlines::*
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see Enums\Gridlines::*
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#abd9af70f706b738690ebb9103940b0e1
     */
    public function gridlines(int $option): self
    {
        return $this->proxy(
            __METHOD__,
            'gridlines',
            false,
            [$option]
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a931114d5266e03931f9fc1b21bd9b56c
     */
    public function centerHorizontally(): self
    {
        return $this->proxy(__METHOD__, 'center_horizontally');
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a3a7fdfe9ef603ef75a03818e6abbfd12
     */
    public function centerVertically(): self
    {
        return $this->proxy(__METHOD__, 'center_vertically');
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#adb169a9740135f5c01eab94099d77f9b
     */
    public function printRowColHeaders(): self
    {
        return $this->proxy(__METHOD__, 'print_row_col_headers');
    }

    /**
     * @param int $firstRow
     * @param int $lastRow
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a5022e812b6fe0414f2a8bab03fd0ec48
     */
    public function repeatRows(int $firstRow, int $lastRow): self
    {
        return $this->proxy(
            __METHOD__,
            'repeat_rows',
            true,
            [
                $firstRow,
                $lastRow,
            ]
        );
    }

    /**
     * @param int[]|string $cols
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a932dd78e18c8e82a21f87e5344aa9c65
     */
    public function repeatColumns($cols): self
    {
        return $this->proxy(
            __METHOD__,
            'repeat_columns',
            true,
            FFILibXlsxWriter::cols($cols)
        );
    }

    /**
     * @param int[]|string $range
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#ab92116759b56cd67419febf6e0aa2c9e
     */
    public function printArea($range): self
    {
        return $this->proxy(
            __METHOD__,
            'print_area',
            true,
            FFILibXlsxWriter::range($range)
        );
    }

    /**
     * @param int $width
     * @param int $height
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9ed8b0603df10ce2439b54f7980efb6d
     */
    public function fitToPages(int $width, int $height): self
    {
        return $this->proxy(
            __METHOD__,
            'fit_to_pages',
            false,
            [$width, $height]
        );
    }

    /**
     * @param int $startPage
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a817235e5ce7b1ea030bedd86ddbb9844
     */
    public function setStartPage(int $startPage): self
    {
        return $this->proxy(
            __METHOD__,
            'set_start_page',
            false,
            [$startPage]
        );
    }

    /**
     * @param int $scale
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a26130cfccecf80a9833d65a84039374d
     */
    public function setPrintScale(int $scale): self
    {
        return $this->proxy(
            __METHOD__,
            'set_print_scale',
            false,
            [$scale]
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a6e32eb8e5f17aedb048035a5df337538
     */
    public function rightToLeft(): self
    {
        return $this->proxy(__METHOD__, 'right_to_left');
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a0b02d53154397cd4024e8ec2a95e2d89
     */
    public function hideZero(): self
    {
        return $this->proxy(__METHOD__, 'hide_zero');
    }

    /**
     * @param int $color
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a1e84ef4dff791fc2278dfb029af94cb0
     */
    public function setTabColor(int $color): self
    {
        return $this->proxy(
            __METHOD__,
            'set_tab_color',
            false,
            [$color]
        );
    }

    /**
     * @param string $password
     * @param Protection|null $protection
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a1b49e135d4debcdb25941f2f40f04cb0
     */
    public function protect(string $password = null, Protection $protection = null): self
    {
        return $this->proxy(
            __METHOD__,
            'protect',
            false,
            [
                $password,
                $protection ? $protection->getPointer() : null
            ]
        );
    }

    /**
     * @param bool $visible
     * @param bool $below
     * @param bool $right
     * @param bool $autoStyle
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a9efae5027e762c9a29f6afa547e4b2db
     */
    public function worksheetOutlineSettings(bool $visible, bool $below, bool $right, bool $autoStyle): self
    {
        return $this->proxy(
            __METHOD__,
            'outline_settings',
            false,
            [
                $visible,
                $below,
                $right,
                $autoStyle,
            ]
        );
    }

    /**
     * @param float $height
     * @param bool $hideUnusedRows
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a74121de7c4f67639b30f302d4af48eb4
     */
    public function worksheetSetDefaultRow(float $height, bool $hideUnusedRows): self
    {
        return $this->proxy(
            __METHOD__,
            'set_default_row',
            false,
            [
                $height,
                $hideUnusedRows
            ]
        );
    }

    /**
     * @param string $name
     * @return Worksheet
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#a8df6c0ce82172dafdd3e81923978bd4b
     */
    public function worksheetSetVbaName(string $name): self
    {
        return $this->proxy(
            __METHOD__,
            'set_vba_name',
            true,
            [$name]
        );
    }

    /**
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#aafb945c6f7f462e7c6eb24032ea4e61a
     */
    public function worksheetShowComments(): self
    {
        return $this->proxy(__METHOD__, 'show_comments');
    }

    /**
     * @param string $author
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/worksheet_8h.html#af196b0cb611f3abc800f3439f3bb8942
     */
    public function worksheetSetCommentsAuthor(string $author): self
    {
        return $this->proxy(
            __METHOD__,
            'set_comments_author',
            false,
            [$author]
        );
    }
}
