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
    /**
     * RowColOptions constructor.
     */
    public function __construct()
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_row_col_options', false, false);
        $this->pointer = FFI::addr($this->struct);
    }
}
