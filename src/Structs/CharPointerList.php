<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFI\CType;

class CharPointerList extends AbstractList
{
    protected function castPart($part)
    {
        return is_string($part) ? new CharPointer($part) : $part;
    }

    protected function getItemType(): CType
    {
        return FFI::type('char*');
    }

    protected function partToStructItem($part)
    {
        if (!$part instanceof CharPointer) {
            throw new \RuntimeException('invalid part');
        }

        return $part->getStruct();
    }

    protected function shouldAddNullToEnd(): bool
    {
        return true;
    }
}
