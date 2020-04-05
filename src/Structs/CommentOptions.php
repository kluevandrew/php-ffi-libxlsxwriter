<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;

/**
 * Class CommentOptions
 * @property int $visible self::COMMENT_DISPLAY_*
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
    public const COMMENT_DISPLAY_DEFAULT = 0;
    public const COMMENT_DISPLAY_HIDDEN = 1;
    public const COMMENT_DISPLAY_VISIBLE = 2;

    /**
     * CommentOptions constructor.
     */
    public function __construct()
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_comment_options');
        $this->pointer = FFI::addr($this->struct);
    }
}
