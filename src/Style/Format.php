<?php

namespace FFILibXlsxWriter\Style;

use FFI\CData;
use FFILibXlsxWriter\Contracts\Closable;
use FFILibXlsxWriter\Contracts\DoNotFreeDirectly;
use FFILibXlsxWriter\Exceptions\CallAfterCloseException;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;
use FFILibXlsxWriter\Traits\ClosableStruct;
use FFILibXlsxWriter\Workbook;
use FFILibXlsxWriter\Enums;

/**
 * Class Format
 */
class Format extends Struct implements DoNotFreeDirectly, Closable
{
    use ClosableStruct;

    /**
     * @var Workbook
     */
    private Workbook $workbook;

    /**
     * Format constructor.
     * @param Workbook $workbook
     * @param CData|null $pointer
     */
    public function __construct(Workbook $workbook, CData $pointer = null)
    {
        $this->workbook = $workbook;
        $this->pointer = $pointer ?? FFILibXlsxWriter::ffi()->workbook_add_format($this->workbook->getPointer());
        $this->struct = $this->pointer[0];
    }

    /**
     * Set the font used in the cell. More...
     * @param string $fontName
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a449e2235a9088cc60233ae443acd2b1a
     */
    public function setFontName(string $fontName): self
    {
        return $this->proxy(__METHOD__, 'set_font_name', [$fontName]);
    }

    /**
     * Set the size of the font used in the cell. More...
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#aec5c1028fa3e25ad60e439fd64afb245
     */
    public function setFontSize(float $size): self
    {
        return $this->proxy(__METHOD__, 'set_font_size', [$size]);
    }

    /**
     * Set the color of the font used in the cell. More...
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a61981b2080bfe6381ede5358ee05b05c
     */
    public function setFontColor(int $color): self
    {
        return $this->proxy(__METHOD__, 'set_font_color', [$color]);
    }

    /**
     * Turn on bold for the format font. More...
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a85e1b0baf44b445b65894e48722aec14
     */
    public function setBold(): self
    {
        return $this->proxy(__METHOD__, 'set_bold');
    }

    /**
     * Turn on italic for the format font. More...
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a19cbd7c34692eb7fb35a7411432d836e
     */
    public function setItalic(): self
    {
        return $this->proxy(__METHOD__, 'set_italic');
    }

    /**
     * Turn on underline for the format: More...
     * @param int $underline Enums\Underline::*
     * @return $this
     * @throws CallAfterCloseException
     * @see Enums\Underline::*
     * @see http://libxlsxwriter.github.io/format_8h.html#ad35ee5445826bd93ec1bc0d489fc09db
     */
    public function setUnderline(int $underline): self
    {
        return $this->proxy(__METHOD__, 'set_underline', [$underline]);
    }

    /**
     * Set the strikeout property of the font. More...
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#ad6e9600723fd772c3cd4d62599beaf31
     */
    public function setFontStrikeout(): self
    {
        return $this->proxy(__METHOD__, 'set_font_strikeout');
    }

    /**
     * Set the superscript/subscript property of the font. More...
     * @param int $fontScript Enums\FontScript::*
     * @return $this
     * @throws CallAfterCloseException
     * @see Enums\FontScript::*
     * @see http://libxlsxwriter.github.io/format_8h.html#a471ca432e429505c79982ca5aecd1db0
     */
    public function setFontScript(int $fontScript): self
    {
        return $this->proxy(__METHOD__, 'set_font_script', [$fontScript]);
    }

    /**
     * Set the number format for a cell. More...
     * @param string $numFormat
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#af77bbd0003344cb16d455c7fb709e16c
     */
    public function setNumFormat(string $numFormat): self
    {
        return $this->proxy(__METHOD__, 'set_num_format', [$numFormat]);
    }

    /**
     * Set the Excel built-in number format for a cell. More...
     * @param int $index
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a688aa42bcc703d17e125d9a34c721872
     */
    public function setNumFormatIndex(int $index): self
    {
        return $this->proxy(__METHOD__, 'set_num_format_index', [$index]);
    }

    /**
     * Set the cell unlocked state. More...
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a1dfd61b72aab2c28c3d51e53e08df587
     */
    public function setUnlocked(): self
    {
        return $this->proxy(__METHOD__, 'set_unlocked');
    }

    /**
     * Hide formulas in a cell. More...
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a135d94ec48564c997c5a78ca8b8861e2
     */
    public function setHidden(): self
    {
        return $this->proxy(__METHOD__, 'set_hidden');
    }

    /**
     * Set the alignment for data in the cell. More...
     * @param int $alignment Enums\Align::*
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a189c83d1f21b01937f1f730720c33d13
     * @see Enums\Align::*
     */
    public function setAlign(int $alignment): self
    {
        return $this->proxy(__METHOD__, 'set_align', [$alignment]);
    }

    /**
     * Wrap text in a cell. More...
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a56d55dd9257d8f0645c62b296d2c196d
     */
    public function setTextWrap(): self
    {
        return $this->proxy(__METHOD__, 'set_text_wrap');
    }

    /**
     * Set the rotation of the text in a cell. More...
     * @param int $angle
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#ae690004cd77f48646da07796b540c309
     */
    public function setRotation(int $angle): self
    {
        return $this->proxy(__METHOD__, 'set_rotation', [$angle]);
    }

    /**
     * Set the cell text indentation level. More...
     * @param int $level
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a99aea699cd7bb3c56a515c9c9e0caa69
     */
    public function setIndent(int $level): self
    {
        return $this->proxy(__METHOD__, 'set_indent', [$level]);
    }

    /**
     * Turn on the text "shrink to fit" for a cell. More...
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a8fc47dd0e47020358c79e20039cbd760
     */
    public function setShrink(): self
    {
        return $this->proxy(__METHOD__, 'set_shrink');
    }

    /**
     * Set the background fill pattern for a cell. More...
     * @param int $index Enums\Pattern::*
     * @return $this
     * @throws CallAfterCloseException
     * @see Enums\Pattern::*
     * @see http://libxlsxwriter.github.io/format_8h.html#a43ddbc77d637b04fdfbc45e96857d15a
     */
    public function setPattern(int $index): self
    {
        return $this->proxy(__METHOD__, 'set_pattern', [$index]);
    }

    /**
     * Set the pattern background color for a cell. More...
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#aeef47436c335daf1801683ac7b3b587d
     */
    public function setBgColor(int $color): self
    {
        return $this->proxy(__METHOD__, 'set_bg_color', [$color]);
    }

    /**
     * Set the pattern foreground color for a cell. More...
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a65086b2b6ee51fd34893e3c53e0578eb
     */
    public function setFgColor(int $color): self
    {
        return $this->proxy(__METHOD__, 'set_fg_color', [$color]);
    }

    /**
     * Set the cell border style. More...
     * @param int $style Enums\Border::*
     * @return $this
     * @throws CallAfterCloseException
     * @see Enums\Border::*
     * @see http://libxlsxwriter.github.io/format_8h.html#a9cf7a28a6e8014cb98dff27415e2b1ca
     */
    public function setBorder(int $style): self
    {
        return $this->proxy(__METHOD__, 'set_border', [$style]);
    }

    /**
     * Set the cell bottom border style. More...
     * @param int $style Enums\Border::*
     * @return $this
     * @throws CallAfterCloseException
     * @see Enums\Border::*
     * @see http://libxlsxwriter.github.io/format_8h.html#a05edc61c138b3ba56727efa24592e990
     */
    public function setBottom(int $style): self
    {
        return $this->proxy(__METHOD__, 'set_bottom', [$style]);
    }

    /**
     * Set the cell top border style. More...
     * @param int $style Enums\Border::*
     * @return $this
     * @throws CallAfterCloseException
     * @see Enums\Border::*
     * @see http://libxlsxwriter.github.io/format_8h.html#a39589314f295cf5610a759d233d1e9c5
     */
    public function setTop(int $style): self
    {
        return $this->proxy(__METHOD__, 'set_top', [$style]);
    }

    /**
     * Set the cell left border style. More...
     * @param int $style Enums\Border::*
     * @return $this
     * @throws CallAfterCloseException
     * @see Enums\Border::*
     * @see http://libxlsxwriter.github.io/format_8h.html#a21f80d92069d1c0a422daa954c4c6eaa
     */
    public function setLeft(int $style): self
    {
        return $this->proxy(__METHOD__, 'set_left', [$style]);
    }

    /**
     * Set the cell right border style. More...
     * @param int $style Enums\Border::*
     * @return $this
     * @throws CallAfterCloseException
     * @see Enums\Border::*
     * @see http://libxlsxwriter.github.io/format_8h.html#a4deaaa289159778326c8eb901c70fbb9
     */
    public function setRight(int $style): self
    {
        return $this->proxy(__METHOD__, 'set_right', [$style]);
    }

    /**
     * Set the color of the cell border. More...
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#ad8fa6d2b638012fc6e331fcd5cf4266b
     */
    public function setBorderColor(int $color): self
    {
        return $this->proxy(__METHOD__, 'set_border_color', [$color]);
    }

    /**
     * Set the color of the bottom cell border. More...
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a53d5df0f55f154b1019e19f7db3f7df3
     */
    public function setBottomColor(int $color): self
    {
        return $this->proxy(__METHOD__, 'set_bottom_color', [$color]);
    }

    /**
     * Set the color of the top cell border. More...
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#af1126cbf0f5d4a5832d251572566335e
     */
    public function setTopColor(int $color): self
    {
        return $this->proxy(__METHOD__, 'set_top_color', [$color]);
    }

    /**
     * Set the color of the left cell border. More...
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a72ae1cd4040cc5d8b6c7b10697fe982a
     */
    public function setLeftColor(int $color): self
    {
        return $this->proxy(__METHOD__, 'set_left_color', [$color]);
    }

    /**
     * Set the color of the right cell border. More...
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a8b1965f2420d7803b6ad5d5b33ce73a9
     */
    public function setRightColor(int $color): self
    {
        return $this->proxy(__METHOD__, 'set_right_color', [$color]);
    }

    /**
     * @param string $phpMethod
     * @param string $cMethod
     * @param array $arguments
     * @return $this
     * @throws CallAfterCloseException
     */
    protected function proxy(string $phpMethod, string $cMethod, array $arguments = []): self
    {
        $this->throwIfClosed($phpMethod);

        $cMethod = 'format_' . $cMethod;
        FFILibXlsxWriter::ffi()->{$cMethod}(
            $this->getPointer(),
            ...$arguments
        );

        return $this;
    }
}
