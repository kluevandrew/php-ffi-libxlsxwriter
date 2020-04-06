<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;

/**
 * Class WorkbookOptions
 * @property bool $constantMemory
 * @property bool $useZip64
 * @property CharPointer|string|null $tmpDir
 */
class WorkbookOptions extends Struct
{
    /**
     * LXWDateTime constructor.
     * @param bool $constantMemory
     * @param bool $useZip64
     * @param string|null $tmpDir
     */
    public function __construct(bool $constantMemory = false, bool $useZip64 = false, string $tmpDir = null)
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_workbook_options', false, false);
        $this->pointer = FFI::addr($this->struct);

        $this->constant_memory = $constantMemory;
        $this->use_zip64 = $useZip64;
        $this->tmpdir = $tmpDir;
    }

    protected function getStructuralProperties(): array
    {
        return [
            'tmpdir' => CharPointer::class,
        ];
    }
}
