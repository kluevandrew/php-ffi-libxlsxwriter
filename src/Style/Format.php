<?php

namespace FFILibXlsxWriter\Style;

use FFI\CData;
use FFILibXlsxWriter\Contracts\DoNotFreeDirectly;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;
use FFILibXlsxWriter\Workbook;

/**
 * Class Format
 */
class Format extends Struct implements DoNotFreeDirectly
{
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

    public function free(): void
    {
        $this->pointer = null;
        $this->struct = null;

        parent::free();
    }

    /**
     * @return CData
     * @deprecated
     */
    public function getCFormat(): CData
    {
        return $this->getPointer();
    }

    /**
     * @return $this
     *
     * @see http://libxlsxwriter.github.io/format_8h.html#a85e1b0baf44b445b65894e48722aec14
     */
    public function setBold(): self
    {
        FFILibXlsxWriter::ffi()->format_set_bold($this->getPointer());

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
        FFILibXlsxWriter::ffi()->format_set_num_format($this->getPointer(), $numFormat);

        return $this;
    }

    public function setItalic(): self
    {
        FFILibXlsxWriter::ffi()->format_set_italic($this->getPointer());

        return $this;
    }

    public function setFontColor(int $color): self
    {
        FFILibXlsxWriter::ffi()->format_set_font_color($this->getPointer(), $color);

        return $this;
    }

    public function setAlign(int $align): self
    {
        FFILibXlsxWriter::ffi()->format_set_align($this->getPointer(), $align);

        return $this;
    }

    public function setFontScript(int $fontScript): self
    {
        FFILibXlsxWriter::ffi()->format_set_font_script($this->getPointer(), $fontScript);

        return $this;
    }

    public function setUnderline(int $underline): self
    {
        FFILibXlsxWriter::ffi()->format_set_underline($this->getPointer(), $underline);

        return $this;
    }

    public function setBgColor(int $color): self
    {
        FFILibXlsxWriter::ffi()->format_set_bg_color($this->getPointer(), $color);

        return $this;
    }

    /**
     * @param int $color
     * @return $this
     */
    public function setFgColor(int $color): self
    {
        FFILibXlsxWriter::ffi()->format_set_fg_color($this->getPointer(), $color);

        return $this;
    }

    public function setBorder(int $border): self
    {
        FFILibXlsxWriter::ffi()->format_set_border($this->getPointer(), $border);

        return $this;
    }

    public function setUnlocked(): self
    {
        FFILibXlsxWriter::ffi()->format_set_unlocked($this->getPointer());

        return $this;
    }

    public function setHidden(): self
    {
        FFILibXlsxWriter::ffi()->format_set_hidden($this->getPointer());

        return $this;
    }

    /**
     * @return $this
     * @see http://libxlsxwriter.github.io/format_8h.html#a56d55dd9257d8f0645c62b296d2c196d
     */
    public function setTextWrap(): self
    {
        FFILibXlsxWriter::ffi()->format_set_text_wrap($this->getPointer());

        return $this;
    }

    /**
     * @param int $level
     * @return $this
     * @see http://libxlsxwriter.github.io/format_8h.html#a99aea699cd7bb3c56a515c9c9e0caa69
     */
    public function setIndent(int $level): self
    {
        FFILibXlsxWriter::ffi()->format_set_indent($this->getPointer(), $level);

        return $this;
    }
}
