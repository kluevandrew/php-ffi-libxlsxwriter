<?php

namespace FFILibXlsxWriter\Options;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;
use FFILibXlsxWriter\Structs\CharPointer;

/**
 * Class ImageOptions
 * @property int $x_offset
 * @property int $y_offset
 * @property float $x_scale
 * @property float $y_scale
 * @property int $object_position ObjectPosition::*
 * @property CharPointer|string $description
 * @property CharPointer|string $url
 * @property CharPointer|string $tip
 */
class ImageOptions extends Struct
{
    /**
     * ImageOptions constructor.
     */
    public function __construct()
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_image_options', false, false);
        $this->pointer = FFI::addr($this->struct);
    }

    /**
     * @return array
     */
    protected function getStructuralProperties(): array
    {
        return [
            'description' => CharPointer::class,
            'url' => CharPointer::class,
            'tip' => CharPointer::class,
        ];
    }
}
