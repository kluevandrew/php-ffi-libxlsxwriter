<?php

namespace FFILibXlsxWriter\Options;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;
use FFILibXlsxWriter\Structs\CharPointer;

/**
 * Class DocProperties
 * @property CharPointer|string|null $title
 * @property CharPointer|string|null $subject
 * @property CharPointer|string|null $author
 * @property CharPointer|string|null $manager
 * @property CharPointer|string|null $company
 * @property CharPointer|string|null $category
 * @property CharPointer|string|null $keywords
 * @property CharPointer|string|null $comments
 * @property CharPointer|string|null $status
 * @property CharPointer|string|null $hyperlink_base
 * @property int|null $created
 */
class DocProperties extends Struct
{
    /**
     * CommentOptions constructor.
     */
    public function __construct()
    {
        $ffi = FFILibXlsxWriter::ffi();
        $this->struct = $ffi->new('struct lxw_doc_properties', false, false);
        $this->pointer = FFI::addr($this->struct);
    }

    protected function getStructuralProperties(): array
    {
        return [
            'title' => CharPointer::class,
            'subject' => CharPointer::class,
            'author' => CharPointer::class,
            'manager' => CharPointer::class,
            'company' => CharPointer::class,
            'category' => CharPointer::class,
            'keywords' => CharPointer::class,
            'comments' => CharPointer::class,
            'status' => CharPointer::class,
            'hyperlink_base' => CharPointer::class,
        ];
    }
}
