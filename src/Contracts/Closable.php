<?php

namespace FFILibXlsxWriter\Contracts;

interface Closable
{
    public function close(): void;
    public function isClosed(): bool;
}
