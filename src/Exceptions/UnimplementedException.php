<?php

namespace FFILibXlsxWriter\Exceptions;

use Throwable;

class UnimplementedException extends FFILibXlsxWriterException
{
    public function __construct(string $method, Throwable $previous = null)
    {
        parent::__construct(sprintf('Method %s is not implemented yet', $method), 100, $previous);
    }
}
