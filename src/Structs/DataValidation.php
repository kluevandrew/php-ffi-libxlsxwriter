<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;

/**
 * Class DataValidation
 * @property int $validate ValidationType::*
 * @property int $criteria ValidationCriteria::*
 * @property bool $ignore_blank
 * @property bool $show_input
 * @property bool $show_error
 * @property int $error_type ValidationError::*
 * @property bool $dropdown
 * @property double $value_number
 * @property CharPointer|string|null $value_formula
 * @property CharPointerList $value_list (char **)
 * @property DateTime $value_datetime
 * @property double $minimum_number
 * @property CharPointer|string|null $minimum_formula
 * @property DateTime $minimum_datetime
 * @property double $maximum_number
 * @property CharPointer|string|null $maximum_formula
 * @property DateTime $maximum_datetime
 * @property CharPointer|string|null $input_title
 * @property CharPointer|string|null $input_message
 * @property CharPointer|string|null $error_title
 * @property CharPointer|string|null $error_message
 */
class DataValidation extends Struct
{
    /**
     * DataValidation constructor.
     */
    public function __construct()
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_data_validation', false, false);
        $this->pointer = FFI::addr($this->struct);
    }

    protected function getStructuralProperties(): array
    {
        return [
            'value_formula' => CharPointer::class,
            'minimum_formula' => CharPointer::class,
            'maximum_formula' => CharPointer::class,
            'input_title' => CharPointer::class,
            'input_message' => CharPointer::class,
            'error_title' => CharPointer::class,
            'error_message' => CharPointer::class,
            'value_list' => CharPointerList::class,
            'value_datetime' => DateTime::class,
            'minimum_datetime' => DateTime::class,
            'maximum_datetime' => DateTime::class,
        ];
    }
}
