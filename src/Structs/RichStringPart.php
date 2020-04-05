<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFI\CData;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Style\Format;

class RichStringPart
{
    /**
     * @var string
     */
    protected string $string;

    /**
     * @var Format|null
     */
    protected ?Format $format;

    /**
     * @var CData
     */
    protected ?CData $struct = null;

    /**
     * @var CData|null
     */
    protected ?CData $pointer = null;

    /**
     * @var CharPointer|null
     */
    protected ?CharPointer $charPointer = null;

    public function __construct(string $string, Format $format = null)
    {
        $this->string = $string;
        $this->format = $format;
    }

    private function allocate(): void
    {
        $ffi = FFILibXlsxWriter::ffi();

        $this->charPointer = new CharPointer($this->string);
        $this->struct = $ffi->new($ffi->type('lxw_rich_string_tuple'));
        $this->struct->string = $this->charPointer->getPointer();
        if (null !== $this->format) {
            $this->struct->format = $this->format->getCFormat();
        }
        $this->pointer = FFI::addr($this->struct);
    }

    /**
     * @param string $string
     * @return RichStringPart
     */
    public function setString(string $string): self
    {
        $this->string = $string;
        $this->free();

        return $this;
    }

    /**
     * @param Format|null $format
     * @return RichStringPart
     */
    public function setFormat(?Format $format): self
    {
        $this->format = $format;
        if (null !== $this->struct) {
            $this->struct->format = $this->format->getCFormat();
        }

        return $this;
    }

    /**
     * @return CData
     */
    public function getPointer(): CData
    {
        if (null === $this->pointer) {
            $this->allocate();
        }

        return $this->pointer;
    }

    /**
     * Deallocate
     */
    protected function free(): void
    {
        if ($this->charPointer) {
            $this->charPointer->free();
            $this->charPointer = null;
        }

        if ($this->pointer) {
            if (!FFI::isNull($this->pointer)) {
                FFI::free($this->pointer);
            }
            $this->pointer = null;
        }

        $this->struct = null;
    }
}
