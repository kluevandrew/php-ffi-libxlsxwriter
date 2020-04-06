<?php

namespace FFILibXlsxWriter;

use FFILibXlsxWriter\Contracts\Closable;
use FFILibXlsxWriter\Contracts\DoNotFreeDirectly;
use FFILibXlsxWriter\Exceptions\CallAfterCloseException;
use FFILibXlsxWriter\Options\WorkbookOptions;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Traits\ClosableStruct;

class Workbook extends Struct implements DoNotFreeDirectly, Closable
{
    use ClosableStruct {
        ClosableStruct::close as private traitClose;
    }

    /**
     * @var Format|null
     */
    protected ?Format $defaultUrlFormat = null;

    /**
     * @var bool
     */
    protected bool $closed = false;

    /**
     * Workbook constructor.
     * @param string $path
     * @param WorkbookOptions|null $options
     */
    public function __construct(string $path, WorkbookOptions $options = null)
    {
        if (null !== $options) {
            $this->pointer = FFILibXlsxWriter::ffi()->workbook_new_opt($path, $options->getPointer());
        } else {
            $this->pointer = FFILibXlsxWriter::ffi()->workbook_new($path);
        }
        $this->struct = $this->pointer[0];
    }

    /**
     * @param string|null $name
     * @return Worksheet
     * @throws CallAfterCloseException
     */
    public function addWorksheet(string $name = null): Worksheet
    {
        $this->throwIfClosed(__METHOD__);

        $worksheet = new Worksheet($this, $name);

        $this->addStructure($worksheet);

        return $worksheet;
    }

    /**
     * @return Format
     * @throws CallAfterCloseException
     */
    public function addFormat(): Format
    {
        $this->throwIfClosed(__METHOD__);

        $format = new Format($this);

        $this->addStructure($format);

        return $format;
    }

    /**
     * @return Format
     * @throws CallAfterCloseException
     */
    public function getDefaultUrlFormat(): Format
    {
        $this->throwIfClosed(__METHOD__);

        if (null === $this->defaultUrlFormat) {
            $pointer = FFILibXlsxWriter::ffi()->workbook_get_default_url_format($this->getPointer());
            $this->defaultUrlFormat = new Format($this, $pointer);
        }

        return $this->defaultUrlFormat;
    }

    public function close(): void
    {
        FFILibXlsxWriter::ffi()->workbook_close($this->getPointer());
        $this->traitClose();
    }
}
