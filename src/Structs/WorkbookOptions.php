<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFI\CData;
use FFILibXlsxWriter\FFILibXlsxWriter;

/**
 * Class WorkbookOptions
 * @property bool $constantMemory
 * @property bool $useZip64
 * @property string|null $tmpDir
 */
class WorkbookOptions extends Struct
{
    /**
     * @var CData|null
     */
    private ?CData $tmpDirPointer = null;

    /**
     * LXWDateTime constructor.
     * @param bool $constantMemory
     * @param bool $useZip64
     * @param string|null $tmpDir
     */
    public function __construct(bool $constantMemory = false, bool $useZip64 = false, string $tmpDir = null)
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_workbook_options');
        $this->pointer = FFI::addr($this->struct);

        $this->struct->constant_memory = $constantMemory;
        $this->struct->use_zip64 = $useZip64;
        if ($tmpDir) {
            $this->tmpDirPointer = $ffi->strdup($tmpDir);
            $this->struct->tmpdir = $this->tmpDirPointer;
        }
    }

    /**
     * @inheritDoc
     */
    public function free(): void
    {
        if ($this->tmpDirPointer !== null && !FFI::isNull($this->tmpDirPointer)) {
            FFI::free($this->tmpDirPointer);
            $this->tmpDirPointer = null;
        }

        parent::free();
    }
}
