<?php

namespace FFILibXlsxWriter\Style;

use FFI\CData;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

/**
 * Class Format
 */
class Format
{
    /**
     * @var Workbook
     */
    private Workbook $workbook;
    private CData $cFormat;

    /**
     * Format constructor.
     * @param Workbook $workbook
     * @param CData|null $cFormat
     */
    public function __construct(Workbook $workbook, CData $cFormat = null)
    {
        $this->workbook = $workbook;
        $this->cFormat = $cFormat ?? FFILibXlsxWriter::ffi()->workbook_add_format($this->workbook->getCWorkbook());
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

    public function setItalic(): self
    {
        FFILibXlsxWriter::ffi()->format_set_italic($this->cFormat);

        return $this;
    }

    public function setFontColor(int $color): self
    {
        FFILibXlsxWriter::ffi()->format_set_font_color($this->cFormat, $color);


        return $this;
    }

    public function setAlign(int $align): self
    {
        FFILibXlsxWriter::ffi()->format_set_align($this->cFormat, $align);

        return $this;
    }

    public function setFontScript(int $fontScript): self
    {
        FFILibXlsxWriter::ffi()->format_set_font_script($this->cFormat, $fontScript);

        return $this;
    }

    public function setUnderline(int $underline): self
    {
        FFILibXlsxWriter::ffi()->format_set_underline($this->cFormat, $underline);

        return $this;
    }

    public function setBgColor(int $color): self
    {
        FFILibXlsxWriter::ffi()->format_set_bg_color($this->cFormat, $color);

        return $this;
    }

    /**
     * @param int $color
     * @return $this
     */
    public function setFgColor(int $color): self
    {
        FFILibXlsxWriter::ffi()->format_set_fg_color($this->cFormat, $color);

        return $this;
    }

    public function setBorder(int $border): self
    {
        FFILibXlsxWriter::ffi()->format_set_border($this->cFormat, $border);

        return $this;
    }

    public function setUnlocked(): self
    {
        FFILibXlsxWriter::ffi()->format_set_unlocked($this->cFormat);

        return $this;
    }

    public function setHidden(): self
    {
        FFILibXlsxWriter::ffi()->format_set_hidden($this->cFormat);

        return $this;
    }

}
