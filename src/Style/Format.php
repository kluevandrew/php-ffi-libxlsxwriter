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

    public function free(): void
    {
        $this->pointer = null;
        $this->struct = null;

        parent::free();
    }

    /**
     * @return $this
     *
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a85e1b0baf44b445b65894e48722aec14
     */
    public function setBold(): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_bold($this->getPointer());

        return $this;
    }

    /**
     * @param string $numFormat
     * @return $this
     *
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#af77bbd0003344cb16d455c7fb709e16c
     */
    public function setNumFormat(string $numFormat): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_num_format($this->getPointer(), $numFormat);

        return $this;
    }

    /**
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setItalic(): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_italic($this->getPointer());

        return $this;
    }

    /**
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setFontColor(int $color): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_font_color($this->getPointer(), $color);

        return $this;
    }

    /**
     * @param int $align
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setAlign(int $align): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_align($this->getPointer(), $align);

        return $this;
    }

    /**
     * @param int $fontScript
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setFontScript(int $fontScript): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_font_script($this->getPointer(), $fontScript);

        return $this;
    }

    /**
     * @param int $underline
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setUnderline(int $underline): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_underline($this->getPointer(), $underline);

        return $this;
    }

    /**
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setBgColor(int $color): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_bg_color($this->getPointer(), $color);

        return $this;
    }

    /**
     * @param int $color
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setFgColor(int $color): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_fg_color($this->getPointer(), $color);

        return $this;
    }

    /**
     * @param int $border
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setBorder(int $border): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_border($this->getPointer(), $border);

        return $this;
    }

    /**
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setUnlocked(): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_unlocked($this->getPointer());

        return $this;
    }

    /**
     * @return $this
     * @throws CallAfterCloseException
     */
    public function setHidden(): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_hidden($this->getPointer());

        return $this;
    }

    /**
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a56d55dd9257d8f0645c62b296d2c196d
     */
    public function setTextWrap(): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_text_wrap($this->getPointer());

        return $this;
    }

    /**
     * @param int $level
     * @return $this
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/format_8h.html#a99aea699cd7bb3c56a515c9c9e0caa69
     */
    public function setIndent(int $level): self
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->format_set_indent($this->getPointer(), $level);

        return $this;
    }
}
