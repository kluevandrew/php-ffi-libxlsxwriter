<?php

namespace FFILibXlsxWriter\Exceptions;

class ZeroDataException extends FFILibXlsxWriterException
{
    /**
     * ZeroDataException constructor.
     * @param string $class
     * @param \Throwable|null $previous
     */
    public function __construct(string $class, \Throwable $previous = null)
    {
        parent::__construct(
            sprintf(
                'The %s cannot be instantiated with empty data',
                $class
            ),
            29,
            $previous
        );
    }
}
