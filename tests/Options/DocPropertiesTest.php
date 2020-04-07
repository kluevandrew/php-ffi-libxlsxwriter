<?php

namespace Tests\Options;

use FFILibXlsxWriter\Options\DocProperties;
use FFILibXlsxWriter\Structs\CharPointer;
use Tests\FFITestCase;

class DocPropertiesTest extends FFITestCase
{
    public function testAllProperties(): void
    {
        $docProperties = new DocProperties();

        $chars = [
            'title',
            'subject',
            'author',
            'manager',
            'company',
            'category',
            'keywords',
            'comments',
            'status',
            'hyperlink_base',
        ];

        foreach ($chars as $propName) {
            $docProperties->{$propName} = 'randomString';
            $this->assertInstanceOf(CharPointer::class, $docProperties->{$propName});
        }

        $time = time();
        $docProperties->created = $time;

        $this->assertEquals($time, $docProperties->created);
    }
}
