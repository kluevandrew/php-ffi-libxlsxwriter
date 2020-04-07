<?php

namespace FFILibXlsxWriter;

use FFILibXlsxWriter\Contracts\Closable;
use FFILibXlsxWriter\Contracts\DoNotFreeDirectly;
use FFILibXlsxWriter\Exceptions\CallAfterCloseException;
use FFILibXlsxWriter\Exceptions\FFILibXlsxWriterException;
use FFILibXlsxWriter\Exceptions\UnimplementedException;
use FFILibXlsxWriter\Options\DocProperties;
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
     * @var Worksheet[]
     */
    protected array $worksheets = [];

    /**
     * @var Chartsheet[]
     */
    protected array $chartsheets = [];

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

        $this->worksheets[] = $worksheet;
        $this->addStructure($worksheet);

        return $worksheet;
    }

    /**
     * @param string|null $name
     * @return Chartsheet
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a9cb69c057990a4139f0a19b53b04b805
     */
    public function addChartsheet(string $name = null): Chartsheet
    {
        $this->throwIfClosed(__METHOD__);

        $chartsheet = new Chartsheet($this, $name);

        $this->chartsheets[] = $chartsheet;
        $this->addStructure($chartsheet);

        return $chartsheet;
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
     * @return Workbook
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#aa814fd7f8d2c3ce86a7aa5d5ed127000
     */
    public function setProperties(DocProperties $properties): self
    {
        return $this->proxy(__METHOD__, 'set_properties', [$properties->getPointer()]);
    }

    /**
     * @param string $name
     * @param string $value
     * @return $this
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a337caad1c1cb16fed505a180d77f20d3
     */
    public function setCustomPropertyString(string $name, string $value): self
    {
        return $this->proxy(
            __METHOD__,
            'set_custom_property_string',
            [$name, $value]
        );
    }

    /**
     * @param string $name
     * @param float $value
     * @return Workbook
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#af2d8534bf9d58e48491a761228d31e14
     */
    public function setCustomPropertyNumber(string $name, float $value): self
    {
        return $this->proxy(
            __METHOD__,
            'set_custom_property_number',
            [$name, $value]
        );
    }

    /**
     * @param string $name
     * @param bool $value
     * @return Workbook
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a929d2f7063651d872ccac14180e042a4
     */
    public function setCustomPropertyBoolean(string $name, bool $value): self
    {
        return $this->proxy(
            __METHOD__,
            'set_custom_property_boolean',
            [$name, $value]
        );
    }

    /**
     * @param string $name
     * @param DateTime $value
     * @return Workbook
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a843f0d1132a7eb3f8e3e9518ff27e5c9
     */
    public function setCustomPropertyDateTime(string $name, DateTime $value): self
    {
        return $this->proxy(
            __METHOD__,
            'set_custom_property_datetime',
            [$name, $value->getPointer()]
        );
    }

    /**
     * @param string $name
     * @param string $formula
     * @return Workbook
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a442b4056e8d4debf56c07888f2a776f6
     */
    public function defineName(string $name, string $formula): self
    {
        return $this->proxy(
            __METHOD__,
            'define_name',
            [$name, $formula]
        );
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
     * @return Worksheet|null
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a857f1d02d04a7d5a039642abde91c280
     */
    public function getWorksheetByName(string $name): ?Worksheet
    {
        $this->throwIfClosed(__METHOD__);

        $worksheetPointer = FFILibXlsxWriter::ffi()->workbook_get_worksheet_by_name(
            $this->getPointer(),
            $name
        );
        if ($worksheetPointer === null) {
            return null;
        }

        foreach ($this->worksheets as $worksheet) {
            if ($worksheet->getPointer() == $worksheetPointer) {
                return $worksheet;
            }
        }

        return null;
    }

    /**
     * @param string $name
     * @return Chartsheet|null
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a857f1d02d04a7d5a039642abde91c280
     */
    public function getChartsheetByName(string $name): ?Chartsheet
    {
        $this->throwIfClosed(__METHOD__);

        $chartsheetPointer = FFILibXlsxWriter::ffi()->workbook_get_chartsheet_by_name(
            $this->getPointer(),
            $name
        );
        if ($chartsheetPointer === null) {
            return null;
        }

        foreach ($this->chartsheets as $chartsheet) {
            if ($chartsheet->getPointer() == $chartsheetPointer) {
                return $chartsheet;
            }
        }

        return null;
    }

    /**
     * This function is used to validate a worksheet or chartsheet name according to the rules used by Excel:
     * The name is less than or equal to 31 UTF-8 characters.
     * The name doesn't contain any of the characters: [ ] : * ? / \
     * The name doesn't start or end with an apostrophe.
     * The name isn't "History", which is reserved by Excel. (Case insensitive).
     * The name isn't already in use. (Case insensitive, see the note below).
     *
     * @param string $name
     * @return bool
     * @throws CallAfterCloseException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a9ecbd616a8306c88210794140c434031
     */
    public function validateSheetname(string $name): bool
    {
        $this->throwIfClosed(__METHOD__);

        $result = FFILibXlsxWriter::ffi()->workbook_validate_sheet_name(
            $this->getPointer(),
            $name
        );

        return $result === 0;
    }

    /**
     * @param string $filename
     * @return Workbook
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a8de478eed94be65de0622c64a0179ff9
     */
    public function addVbaProject(string $filename): self
    {
        return $this->proxy(
            __METHOD__,
            'add_vba_project',
            [$filename]
        );
    }

    /**
     * @param string $name
     * @return Workbook
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     * @see http://libxlsxwriter.github.io/workbook_8h.html#a24dd6624ffdc3e17aa5bcfb4608755b8
     */
    public function setVbaName(string $name): self
    {
        return $this->proxy(
            __METHOD__,
            'set_vba_name',
            [$name]
        );
    }

    /**
     * @param string $phpMethod
     * @param string $cMethod
     * @param array $arguments
     * @return $this
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     */
    protected function proxy(string $phpMethod, string $cMethod, array $arguments = []): self
    {
        $this->throwIfClosed($phpMethod);

        $cMethod = 'workbook_' . $cMethod;
        $code = FFILibXlsxWriter::ffi()->{$cMethod}(
            $this->getPointer(),
            ...$arguments
        );

        if ($code !== 0) {
            throw FFILibXlsxWriterException::byCode(FFILibXlsxWriter::getLog(), $code);
        }

        return $this;
    }

    /**
     * PHP segfaults when dump CDATA of workbook after adding first sheet
     * @return array
     * @throws CallAfterCloseException
     */
    public function __debugInfo(): array
    {
        return [
            'closed' => $this->closed,
            'worksheets' => $this->worksheets,
            'structures' => $this->structures,
            'defaultUrlFormat' => $this->getDefaultUrlFormat(),
        ];
    }
}
