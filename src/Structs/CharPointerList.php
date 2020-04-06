<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFI\CData;
use FFILibXlsxWriter\FFILibXlsxWriter;

class CharPointerList extends Struct
{
    /**
     * @var int
     */
    protected int $size;

    /**
     * @var CharPointer[]
     */
    protected array $parts = [];

    public function __construct(...$strings)
    {
        foreach ($strings as $string) {
            $this->parts[] = is_string($string) ? new CharPointer($string) : $string;
        }
    }

    /**
     * @param CharPointer|string $string
     * @return $this
     */
    public function push($string): self
    {
        $this->free();
        $this->parts[] = is_string($string) ? new CharPointer($string) : $string;

        return $this;
    }

    /**
     * @return void
     */
    private function allocate(): void
    {
        $ffi = FFILibXlsxWriter::ffi();

        $size = count($this->parts);
        $this->struct = $ffi->new(
            FFI::arrayType(
                $ffi->type('char*'),
                [$size + 1]
            ),
            false,
            false
        );

        foreach ($this->parts as $i => $part) {
            $this->struct[$i] = $part->getStruct();
        }

        $this->struct[$size] = null;
        $this->pointer = FFI::addr($this->struct);
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
     * @return CData
     */
    public function getStruct(): CData
    {
        if (null === $this->struct) {
            $this->allocate();
        }

        return $this->struct;
    }
}
