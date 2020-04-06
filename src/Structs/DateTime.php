<?php

namespace FFILibXlsxWriter\Structs;

use DateTime as PHPDateTime;
use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;

/**
 * Class DateTime
 * @property int $year
 * @property int $month
 * @property int $day
 * @property int $hour
 * @property int $min
 * @property float $sec
 */
class DateTime extends Struct
{
    /**
     * DateTime constructor.
     * @param int $year
     * @param int $month
     * @param int $day
     * @param int $hour
     * @param int $min
     * @param float $sec
     */
    public function __construct(int $year, int $month, int $day, int $hour = 0, int $min = 0, float $sec = 0)
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_datetime', false, false);
        $this->pointer = FFI::addr($this->struct);

        $this->struct->year = $year;
        $this->struct->month = $month;
        $this->struct->day = $day;
        $this->struct->hour = $hour;
        $this->struct->min = $min;
        $this->struct->sec = $sec;
    }

    public static function fromDateTime(PHPDateTime $dateTime): self
    {
        list($y, $m, $d, $h, $i, $s) = explode('|', $dateTime->format('Y|m|d|H|i|s.v'));

        return new self($y, $m, $d, $h, $i, $s);
    }

    public static function now()
    {
        return self::fromDateTime(new PHPDateTime());
    }
}
