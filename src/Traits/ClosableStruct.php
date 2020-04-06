<?php

namespace FFILibXlsxWriter\Traits;

use FFI\CData;
use FFILibXlsxWriter\Contracts\Closable;
use FFILibXlsxWriter\Exceptions\CallAfterCloseException;

trait ClosableStruct
{
    protected bool $closed = false;

    abstract protected function free(): void;

    abstract public function getPointer(): ?CData;

    public function isClosed(): bool
    {
        return $this->closed;
    }

    /**
     * Save file
     * @throws CallAfterCloseException
     */
    public function close(): void
    {
        foreach ($this->getStructures() as $structure) {
            if ($structure instanceof Closable && !$structure->isClosed()) {
                $structure->close();
            }
        }
        $this->throwIfClosed(__METHOD__);
        $this->free();
        $this->closed = true;
    }

    /**
     * @param string $method
     * @throws CallAfterCloseException
     */
    protected function throwIfClosed(string $method)
    {
        if ($this->isClosed()) {
            throw new CallAfterCloseException($method);
        }
    }
}
