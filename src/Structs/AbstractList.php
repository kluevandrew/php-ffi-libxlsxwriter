<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFI\CData;
use FFI\CType;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;

abstract class AbstractList extends Struct
{
    /**
     * @var int
     */
    protected int $size;

    /**
     * @var array
     */
    protected array $parts = [];

    public function __construct(...$parts)
    {
        foreach ($parts as $part) {
            $this->push($part);
        }
    }

    /**
     * @param mixed $part
     * @return $this
     */
    public function push($part): self
    {
        $this->free();
        $this->parts[] = $this->castPart($part);

        return $this;
    }

    /**
     * @return void
     */
    private function allocate(): void
    {
        $ffi = FFILibXlsxWriter::ffi();

        $size = count($this->parts);
        if ($this->shouldAddNullToEnd()) {
            $size += 1;
        }
        $this->struct = $ffi->new(
            FFI::arrayType(
                $this->getItemType(),
                [$size + 1]
            ),
            false,
            false
        );

        foreach ($this->parts as $i => $part) {
            $this->struct[$i] = $this->partToStructItem($part);
        }

        if ($this->shouldAddNullToEnd()) {
            $this->struct[$size] = null;
        }
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

    public function getParts(): array
    {
        return $this->parts;
    }

    abstract protected function getItemType(): CType;

    abstract protected function partToStructItem($part);

    protected function shouldAddNullToEnd(): bool
    {
        return false;
    }

    protected function castPart($part)
    {
        return $part;
    }
}
