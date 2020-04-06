<?php

namespace FFILibXlsxWriter\Structs;

use FFI\CType;
use FFILibXlsxWriter\FFILibXlsxWriter;

class IntegerList extends AbstractList
{
    /**
     * @var string
     */
    protected string $type;

    /**
     * IntegerList constructor.
     * @param string $type
     * @param int[] $parts
     */
    public function __construct(string $type, ...$parts)
    {
        $this->type = $type;
        parent::__construct(...$parts);
    }

    protected function castPart($part)
    {
        return (int)$part;
    }

    protected function getItemType(): CType
    {
        return FFILibXlsxWriter::ffi()->type($this->type);
    }

    protected function partToStructItem($part)
    {
        return $part;
    }
}
