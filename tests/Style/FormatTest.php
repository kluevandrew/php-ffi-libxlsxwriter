<?php

namespace Tests\Style;

use FFILibXlsxWriter\Enums\Align;
use FFILibXlsxWriter\Enums\Border;
use FFILibXlsxWriter\Enums\Color;
use FFILibXlsxWriter\Enums\Font;
use FFILibXlsxWriter\Enums\Underline;
use FFILibXlsxWriter\Enums\ValidationCriteria;
use FFILibXlsxWriter\Exceptions\CallAfterCloseException;
use FFILibXlsxWriter\Structs\DataValidation;
use FFILibXlsxWriter\Structs\DateTime;
use FFILibXlsxWriter\Structs\ImageBuffer;
use FFILibXlsxWriter\Structs\RichString;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Workbook;
use FFILibXlsxWriter\Worksheet;
use Tests\FFITestCase;

class FormatTest extends FFITestCase
{
    /**
     * @param callable $callable
     * @dataProvider callAfterCloseProvider
     * @throws CallAfterCloseException
     */
    public function testCallAfterClose(callable $callable): void
    {
        $workbook = new Workbook(sys_get_temp_dir() . '/test.xlsx');
        $format = $workbook->addFormat();
//        $reflectionCLass= new \ReflectionClass($format);
//        $methods = $reflectionCLass->getMethods(\ReflectionMethod::IS_PUBLIC);
//        $string = '';
//        foreach ($methods as $method) {
//            $string .= "'{$method->name}' => [fn(Format \$format) => \$format->{$method->name}()],\n";
//        }
//        echo PHP_EOL;
//        echo $string;
//        echo PHP_EOL;
//        die;
        $callable($format);
        $workbook->close();

        $this->expectException(CallAfterCloseException::class);

        $callable($format);
    }

    /**
     * @return array<array<int, callable>>
     */
    public function callAfterCloseProvider(): array
    {
        return [
            'setBold' => [fn(Format $format) => $format->setBold()],
            'setNumFormat' => [fn(Format $format) => $format->setNumFormat('0.00')],
            'setItalic' => [fn(Format $format) => $format->setItalic()],
            'setFontColor' => [fn(Format $format) => $format->setFontColor(Color::RED)],
            'setAlign' => [fn(Format $format) => $format->setAlign(Align::VERTICAL_CENTER)],
            'setFontScript' => [fn(Format $format) => $format->setFontScript(Font::SUPERSCRIPT)],
            'setUnderline' => [fn(Format $format) => $format->setUnderline(Underline::DOUBLE)],
            'setBgColor' => [fn(Format $format) => $format->setBgColor(Color::RED)],
            'setFgColor' => [fn(Format $format) => $format->setFgColor(Color::BLUE)],
            'setBorder' => [fn(Format $format) => $format->setBorder(Border::DASH_DOT)],
            'setUnlocked' => [fn(Format $format) => $format->setUnlocked()],
            'setHidden' => [fn(Format $format) => $format->setHidden()],
            'setTextWrap' => [fn(Format $format) => $format->setTextWrap()],
            'setIndent' => [fn(Format $format) => $format->setIndent(2)],
        ];
    }
}
