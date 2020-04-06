<?php

namespace FFILibXlsxWriter;

use FFILibXlsxWriter\Contracts\Closable;
use FFILibXlsxWriter\Contracts\DoNotFreeDirectly;
use FFILibXlsxWriter\Exceptions\CallAfterCloseException;
use FFILibXlsxWriter\Exceptions\FFILibXlsxWriterException;
use FFILibXlsxWriter\Exceptions\UnimplementedException;
use FFILibXlsxWriter\Options\WorkbookOptions;
use FFILibXlsxWriter\Structs\DateTime;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Traits\ClosableStruct;

/**
 * Class Workbook
 * @see http://libxlsxwriter.github.io/workbook_8h.html
 */
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
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a1cf96608a23ee4eb0e8467c15240d00b
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a8ca9bd8c30c618b81ca6180f78b03323
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
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a81d456b4f65a464e78e4a0030ecc3c2e
     */
    public function addWorksheet(string $name = null): Worksheet
    {
        $this->throwIfClosed(__METHOD__);

        $worksheet = new Worksheet($this, $name);

        $this->addStructure($worksheet);

        return $worksheet;
    }

    /**
     * @param string|null $name
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a9cb69c057990a4139f0a19b53b04b805
     */
    public function addChartsheet(string $name = null): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @return Format
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a279a5d7075d09a3931aae9782afede33
     */
    public function addFormat(): Format
    {
        $this->throwIfClosed(__METHOD__);

        $format = new Format($this);

        $this->addStructure($format);

        return $format;
    }

    /**
     * @param int $type
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a63b001ecefdbc4417986a3e344657726
     */
    public function addChart(int $type): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#ad9e7aeebc0fd43562db5bcd527b2ee5e
     */
    public function close(): void
    {
        $this->throwIfClosed(__METHOD__);
        FFILibXlsxWriter::ffi()->workbook_close($this->getPointer());
        $this->traitClose();
    }

    /**
     * @param $properties
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#aa814fd7f8d2c3ce86a7aa5d5ed127000
     */
    public function setProperties($properties): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param string $name
     * @param string $value
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a337caad1c1cb16fed505a180d77f20d3
     */
    public function setCustomPropertyString(string $name, string $value): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param string $name
     * @param float $value
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#af2d8534bf9d58e48491a761228d31e14
     */
    public function setCustomPropertyNumber(string $name, float $value): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param string $name
     * @param bool $value
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a929d2f7063651d872ccac14180e042a4
     */
    public function setCustomPropertyBoolean(string $name, bool $value): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param string $name
     * @param DateTime $value
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a843f0d1132a7eb3f8e3e9518ff27e5c9
     */
    public function setCustomPropertyDateTime(string $name, DateTime $value): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param string $name
     * @param string $formula
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a442b4056e8d4debf56c07888f2a776f6
     */
    public function defineName(string $name, string $formula): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @return Format
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#abd18b07856137e94b34c9aca7e0831b5
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

    /**
     * @param string $name
     * @return Worksheet
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a857f1d02d04a7d5a039642abde91c280
     */
    public function getWorksheetByName(string $name): Worksheet
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param string $name
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a857f1d02d04a7d5a039642abde91c280
     */
    public function getChartsheetByName(string $name): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param string $name
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a9ecbd616a8306c88210794140c434031
     */
    public function validateSheetname(string $name): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param string $filename
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a8de478eed94be65de0622c64a0179ff9
     */
    public function addVbaProject(string $filename): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }

    /**
     * @param string $name
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a24dd6624ffdc3e17aa5bcfb4608755b8
     */
    public function setVbaName(string $name): void
    {
        $this->throwIfClosed(__METHOD__);

        throw new UnimplementedException(__METHOD__);
    }
}
