<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;

/**
 * Class ImageOptions
 * @property int $x_offset
 * @property int $y_offset
 * @property float $x_scale
 * @property float $y_scale
 * @property int $object_position ObjectPosition::*
 * @property string $description
 * @property string $url
 * @property string $tip
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
    protected function getCharPointerProperties(): array
    {
        return ['description', 'url', 'tip'];
    }
}
