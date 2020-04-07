<?php

namespace FFILibXlsxWriter\Options;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;

/**
 * Class RowColOptions
 * @property int $hidden
 * @property int $level
 * @property int $collapsed
 */
class RowColOptions extends Struct
{
    public const DEF_ROW_HEIGHT = 15.0;

    /**
     * RowColOptions constructor.
     * @param int $hidden
     * @param int $level
     * @param int $collapsed
     */
    public function __construct(int $hidden, int $level, int $collapsed)
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_row_col_options', false, false);
        $this->pointer = FFI::addr($this->struct);

        $this->hidden = $hidden;
        $this->level = $level;
        $this->collapsed = $collapsed;
    }
}
