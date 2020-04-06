<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFI\CData;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;
use FFILibXlsxWriter\Style\Format;

/**
 * Class RichStringPart
 * @property CharPointer|string|null $string
 * @property Format $format
 */
class RichStringPart extends Struct
{
    /**
     * @var Format|null
     */
    protected ?Format $phpFormat;

    /**
     * @var CData
     */
    protected ?CData $struct = null;

    /**
     * @var CData|null
     */
    protected ?CData $pointer = null;

    public function __construct(string $string, Format $format = null)
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new($ffi->type('struct lxw_rich_string_tuple'), false, false);
        $this->pointer = FFI::addr($this->struct);

        $this->string = $string;
        $this->format = $format ? $format->getPointer() : null;
    }

    public function setFormatAttribute(?Format $format): void
    {
        $this->phpFormat = $format ? $format->getPointer() : null;
        $this->struct->format = $format ? $format->getPointer() : null;
    }

    public function getFormatAttribute(): ?Format
    {
        return $this->phpFormat;
    }

    protected function getStructuralProperties(): array
    {
        return [
            'string' => CharPointer::class,
        ];
    }
}
