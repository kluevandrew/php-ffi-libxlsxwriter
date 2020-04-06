<?php

namespace FFILibXlsxWriter;

use FFI\CData;
use FFILibXlsxWriter\Structs\WorkbookOptions;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Structs\Struct;

class Workbook
{
    private CData $cWorkbook;

    /**
     * @var Struct[]
     */
    protected array $internalStructs = [];

    /**
     * @var Format|null
     */
    protected ?Format $defaultUrlFormat = null;

    /**
     * @var bool
     */
    protected bool $isClosed = false;

    /**
     * Workbook constructor.
     * @param string $path
     * @param WorkbookOptions|null $options
     */
    public function __construct(string $path, WorkbookOptions $options = null)
    {
        if (null !== $options) {
            $this->cWorkbook = FFILibXlsxWriter::ffi()->workbook_new_opt($path, $options->getPointer());
        } else {
            $this->cWorkbook = FFILibXlsxWriter::ffi()->workbook_new($path);
        }
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
        if ($this->isClosed) {
            throw new \RuntimeException('Already closed');
        }
        $this->isClosed = true;

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
     * @internal
     * @param Struct $struct
     */
    public function addStructure(Struct $struct)
    {
        $this->internalStructs[] = $struct;
    }

    /**
     * @return Format
     */
    public function getDefaultUrlFormat(): Format
    {
        if (null === $this->defaultUrlFormat) {
            $cFormat = FFILibXlsxWriter::ffi()->workbook_get_default_url_format($this->cWorkbook);
            $this->defaultUrlFormat = new Format($this, $cFormat);
        }

        return $this->defaultUrlFormat;
    }
}
