<?php

namespace FFILibXlsxWriter;

use FFI\CData;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Types\LXWStruct;

class Workbook
{
    private CData $cWorkbook;

    /**
     * @var LXWStruct[]
     */
    protected array $internalStructs = [];

    /**
     * Workbook constructor.
     * @param string $path
     */
    public function __construct(string $path)
    {
        $this->cWorkbook = FFILibXlsxWriter::ffi()->workbook_new($path);
    }

    /**
     * @return CData
     */
    public function getCWorkbook(): CData
    {
        return $this->cWorkbook;
    }

    /**
     * @param string|null $name
     * @return Worksheet
     */
    public function addWorksheet(string $name = null): Worksheet
    {
        return new Worksheet($this, $name);
    }

    /**
     * @return Format
     */
    public function addFormat(): Format
    {
        return new Format($this);
    }

    /**
     * Save file
     */
    public function close(): void
    {
        FFILibXlsxWriter::ffi()->workbook_close($this->cWorkbook);
        $this->free();
    }

    /**
     * Clean memory
     */
    protected function free(): void
    {
        foreach ($this->internalStructs as $struct) {
            $struct->free();
        }
    }

    /**
     * @param LXWStruct $struct
     */
    public function addStructure(LXWStruct $struct)
    {
        $this->internalStructs[] = $struct;
    }

}
