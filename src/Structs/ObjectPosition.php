<?php

namespace FFILibXlsxWriter\Structs;

abstract class ObjectPosition
{
    public const DEFAULT = 0;
    public const MOVE_AND_SIZE = 1;
    public const MOVE_DONT_SIZE = 2;
    public const DONT_MOVE_DONT_SIZE = 3;
    public const MOVE_AND_SIZE_AFTER = 4;
}
