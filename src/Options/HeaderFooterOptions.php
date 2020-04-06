<?php

namespace FFILibXlsxWriter\Options;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;

/**
 * Class HeaderFooterOptions
 * @property float $margin
 */
class HeaderFooterOptions extends Struct
{
    /**
     * RowColOptions constructor.
     * @param float $margin
     */
    public function __construct(float $margin = 0.3)
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_header_footer_options', false, false);
        $this->pointer = FFI::addr($this->struct);
        $this->margin = $margin;
    }
}
