<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFI\CData;
use FFILibXlsxWriter\FFILibXlsxWriter;

class RichString
{
    /**
     * @var RichStringPart[]
     */
    protected array $parts = [];

    /**
     * @var CData
     */
    protected ?CData $pointer = null;

    /**
     * RichString constructor.
     * @param RichStringPart[] $parts
     */
    public function __construct(array $parts)
    {
        $this->parts = $parts;
    }

    /**
     * @param RichStringPart $part
     * @return $this
     */
    public function addPart(RichStringPart $part): self
    {
        $this->parts[] = $part;
        $this->pointer = null;

        return $this;
    }

    /**
     * @return void
     */
    private function allocate(): void
    {
        $ffi = FFILibXlsxWriter::ffi();

        $size = count($this->parts);
        $this->pointer = $ffi->new(
            FFI::arrayType(
                $ffi->type('struct lxw_rich_string_tuple*'),
                [$size + 1]
            )
        );

        foreach ($this->parts as $i => $part) {
            $this->pointer[$i] = $part->getPointer();
        }

        $this->pointer[$size] = null;
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
}
