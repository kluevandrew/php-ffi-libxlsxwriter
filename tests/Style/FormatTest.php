<?php

namespace Tests\Style;

use FFILibXlsxWriter\Enums\Align;
use FFILibXlsxWriter\Enums\Border;
use FFILibXlsxWriter\Enums\Color;
use FFILibXlsxWriter\Enums\FontScript;
use FFILibXlsxWriter\Enums\Pattern;
use FFILibXlsxWriter\Enums\Underline;
use FFILibXlsxWriter\Exceptions\CallAfterCloseException;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Workbook;
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
            'setFontScript' => [fn(Format $format) => $format->setFontScript(FontScript::SUPERSCRIPT)],
            'setUnderline' => [fn(Format $format) => $format->setUnderline(Underline::DOUBLE)],
            'setBgColor' => [fn(Format $format) => $format->setBgColor(Color::RED)],
            'setFgColor' => [fn(Format $format) => $format->setFgColor(Color::BLUE)],
            'setBorder' => [fn(Format $format) => $format->setBorder(Border::DASH_DOT)],
            'setUnlocked' => [fn(Format $format) => $format->setUnlocked()],
            'setHidden' => [fn(Format $format) => $format->setHidden()],
            'setTextWrap' => [fn(Format $format) => $format->setTextWrap()],
            'setIndent' => [fn(Format $format) => $format->setIndent(2)],
            'setFontName' => [fn(Format $format) => $format->setFontName('Arial')],
            'setFontSize' => [fn(Format $format) => $format->setFontSize(22)],
            'setFontStrikeout' => [fn(Format $format) => $format->setFontStrikeout()],
            'setNumFormatIndex' => [fn(Format $format) => $format->setNumFormatIndex(1)],
            'setRotation' => [fn(Format $format) => $format->setRotation(10)],
            'setShrink' => [fn(Format $format) => $format->setShrink()],
            'setPattern' => [fn(Format $format) => $format->setPattern(Pattern::DARK_GRAY)],
            'setBottom' => [fn(Format $format) => $format->setBottom(Border::DOUBLE)],
            'setTop' => [fn(Format $format) => $format->setTop(Border::DASHED)],
            'setLeft' => [fn(Format $format) => $format->setLeft(Border::HAIR)],
            'setRight' => [fn(Format $format) => $format->setRight(Border::THICK)],
            'setBorderColor' => [fn(Format $format) => $format->setBorderColor(Color::BROWN)],
            'setBottomColor' => [fn(Format $format) => $format->setBottomColor(Color::BROWN)],
            'setTopColor' => [fn(Format $format) => $format->setTopColor(Color::BROWN)],
            'setLeftColor' => [fn(Format $format) => $format->setLeftColor(Color::BROWN)],
            'setRightColor' => [fn(Format $format) => $format->setRightColor(Color::BROWN)],
        ];
    }
}
