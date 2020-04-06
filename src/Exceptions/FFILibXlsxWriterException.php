<?php

namespace FFILibXlsxWriter\Exceptions;

use Exception;

class FFILibXlsxWriterException extends Exception
{
    /**
     * @param string $log
     * @param int $code
     * @param \Throwable|null $previous
     * @return self
     * @todo implement all error codes
     */
    final public static function byCode(string $log, int $code, \Throwable $previous = null): self
    {
        switch ($code) {
            case 13:
                return new DataValidationException($log, $code, $previous);
        }

        return new self($log, $code, $previous);
    }
}
