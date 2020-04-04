<?php

namespace FFILibXlsxWriter\Types;

use DateTime;
use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;

class LXWDateTime extends LXWStruct
{
    /**
     * LXWDateTime constructor.
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

        $this->struct->hour = $hour;
        $this->struct->min = $min;
        $this->struct->sec = $sec;

        $this->struct->year = $year;
        $this->struct->month = $month;
        $this->struct->day = $day;
        $this->struct->hour = $hour;
        $this->struct->min = $min;
        $this->struct->sec = $sec;
    }

    public static function fromDateTime(DateTime $dateTime): self
    {
        list($y, $m, $d, $h, $i, $s) = explode('|', $dateTime->format('Y|m|d|H|i|s.v'));

        return new self($y, $m, $d, $h, $i, $s);
    }

}
