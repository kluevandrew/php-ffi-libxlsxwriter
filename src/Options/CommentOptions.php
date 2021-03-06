<?php

namespace FFILibXlsxWriter\Options;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;
use FFILibXlsxWriter\Structs\CharPointer;

/**
 * Class CommentOptions
 * @property int $visible CommentDisplay::*
 * @property string $author
 * @property int $width
 * @property int $height
 * @property double $x_scale
 * @property double $y_scale
 * @property int $color
 * @property string $font_name
 * @property double $font_size
 * @property int $font_family
 * @property int $start_row
 * @property int $start_col
 * @property int $x_offset
 * @property int $y_offset
 */
class CommentOptions extends Struct
{
    /**
     * CommentOptions constructor.
     */
    public function __construct()
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_comment_options', false, false);
        $this->pointer = FFI::addr($this->struct);
    }

    protected function getStructuralProperties(): array
    {
        return [
            'author' => CharPointer::class,
            'font_name' => CharPointer::class,
        ];
    }
}
