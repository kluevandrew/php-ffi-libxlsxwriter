<?php

namespace FFILibXlsxWriter\Exceptions;

use Throwable;

class DataValidationException extends FFILibXlsxWriterException
{
    public function __construct(string $log = "", int $code = 0, Throwable $previous = null)
    {
        parent::__construct($this->parseLog($log), $code, $previous);
    }

    /**
     * @param string $log
     * @return string
     * @todo change C messages to PHP message
     */
    private function parseLog(string $log)
    {
        return trim(str_replace('[WARNING]: ', '', $log));
    }
}
