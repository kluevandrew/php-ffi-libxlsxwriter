<?php

namespace FFILibXlsxWriter\Structs;

use FFI\CType;
use FFILibXlsxWriter\FFILibXlsxWriter;

class RichString extends AbstractList
{
    protected function castPart($part)
    {
        return is_string($part) ? new RichStringPart($part) : $part;
    }

    protected function getItemType(): CType
    {
        return FFILibXlsxWriter::ffi()->type('struct lxw_rich_string_tuple*');
    }

    protected function partToStructItem($part)
    {
        if (!$part instanceof RichStringPart) {
            throw new \RuntimeException('invalid part');
        }

        return $part->getStruct();
    }

    protected function shouldAddNullToEnd(): bool
    {
        return true;
    }
}
