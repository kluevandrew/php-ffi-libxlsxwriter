<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;

/**
 * Class Protection
 * @property bool $no_select_locked_cells
 * @property bool $no_select_unlocked_cells
 * @property bool $format_cells
 * @property bool $format_columns
 * @property bool $format_rows
 * @property bool $insert_columns
 * @property bool $insert_rows
 * @property bool $insert_hyperlinks
 * @property bool $delete_columns
 * @property bool $delete_rows
 * @property bool $sort
 * @property bool $autofilter
 * @property bool $pivot_tables
 * @property bool $scenarios
 * @property bool $objects
 * @property bool $no_content
 * @property bool $no_objects
 */
class Protection extends Struct
{
    /**
     * Protection constructor.
     */
    public function __construct()
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_protection', false, false);
        $this->pointer = FFI::addr($this->struct);
    }
}
