<?php

namespace FFILibXlsxWriter\Enums;

abstract class ValidationCriteria
{
    public const NONE = 0;
    public const BETWEEN = 1;
    public const NOT_BETWEEN = 2;
    public const EQUAL_TO = 3;
    public const NOT_EQUAL_TO = 4;
    public const GREATER_THAN = 5;
    public const LESS_THAN = 6;
    public const GREATER_THAN_OR_EQUAL_TO = 7;
    public const LESS_THAN_OR_EQUAL_TO = 8;
}
