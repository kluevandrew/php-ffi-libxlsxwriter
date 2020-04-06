<?php

namespace FFILibXlsxWriter\Exceptions;

class CallAfterCloseException extends FFILibXlsxWriterException
{
    /**
     * CallAfterCloseException constructor.
     * @param string $method
     * @param \Throwable|null $previous
     */
    public function __construct(string $method, \Throwable $previous = null)
    {
        parent::__construct(
            sprintf(
                'The %s method was called after workbook was closed',
                $method
            ),
            28,
            $previous
        );
    }
}
