<?php

namespace FFILibXlsxWriter;

use FFILibXlsxWriter\Contracts\Closable;
use FFILibXlsxWriter\Contracts\DoNotFreeDirectly;
use FFILibXlsxWriter\Exceptions\FFILibXlsxWriterException;
use FFILibXlsxWriter\Traits\ClosableStruct;

abstract class Sheet extends Struct implements DoNotFreeDirectly, Closable
{
    use ClosableStruct;

    protected Workbook $workbook;

    /**
     * Worksheet constructor.
     * @param Workbook $workbook
     * @param string|null $name
     * @see Workbook::addChartsheet()
     */
    public function __construct(Workbook $workbook, ?string $name)
    {
        $this->workbook = $workbook;
        $constructor = 'workbook_add_' . $this->getType();
        $this->pointer = FFILibXlsxWriter::ffi()->{$constructor}(
            $this->workbook->getPointer(),
            $name
        );
        $this->struct = $this->pointer[0];
    }

    abstract protected function getType(): string;

    /**
     * @param string $phpMethod
     * @param string $cMethod
     * @param bool $catch
     * @param array $arguments
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     */
    protected function proxy(string $phpMethod, string $cMethod, bool $catch = false, array $arguments = []): self
    {
        $this->throwIfClosed($phpMethod);

        $cMethod = $this->getType() . '_' . $cMethod;

        $code = FFILibXlsxWriter::ffi()->{$cMethod}(
            $this->getPointer(),
            ...$arguments
        );

        if ($catch && $code !== 0) {
            throw FFILibXlsxWriterException::byCode(FFILibXlsxWriter::getLog(), $code);
        }

        return $this;
    }

    /**
     * @param string $phpMethod
     * @param string $cMethod
     * @param bool $catch
     * @param array $arguments
     * @param Struct|null $options
     * @return $this
     * @throws Exceptions\CallAfterCloseException
     * @throws FFILibXlsxWriterException
     */
    protected function proxyOpt(
        string $phpMethod,
        string $cMethod,
        bool $catch = false,
        array $arguments = [],
        Struct $options = null
    ): self {
        if (null !== $options) {
            $cMethod .= '_opt';
            $arguments[] = $options->getPointer();
        }

        return $this->proxy($phpMethod, $cMethod, $catch, $arguments);
    }
}
