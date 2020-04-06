<?php

namespace FFILibXlsxWriter\Enums;

abstract class ValidationType
{
    public const NONE = 0;
    public const INTEGER = 1;
    public const INTEGER_FORMULA = 2;
    public const DECIMAL = 3;
    public const DECIMAL_FORMULA = 4;
    public const LIST = 5;
    public const LIST_FORMULA = 6;
    public const DATE = 7;
    public const DATE_FORMULA = 8;
    public const DATE_NUMBER = 9;
    public const TIME = 10;
    public const TIME_FORMULA = 11;
    public const TIME_NUMBER = 12;
    public const LENGTH = 13;
    public const LENGTH_FORMULA = 14;
    public const CUSTOM_FORMULA = 15;
    public const ANY = 16;
}
