<?php

/**
 * @see http://libxlsxwriter.github.io/comments2_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\Enums\Align;
use FFILibXlsxWriter\Enums\Color;
use FFILibXlsxWriter\Enums\CommentDisplay;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Options\CommentOptions;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/comments2.xlsx');

$worksheet1 = $workbook->addWorksheet();
$worksheet2 = $workbook->addWorksheet();
$worksheet3 = $workbook->addWorksheet();
$worksheet4 = $workbook->addWorksheet();
$worksheet5 = $workbook->addWorksheet();
$worksheet6 = $workbook->addWorksheet();
$worksheet7 = $workbook->addWorksheet();
$worksheet8 = $workbook->addWorksheet();
$text_wrap = $workbook->addFormat();
$text_wrap->setTextWrap();
$text_wrap->setAlign(Align::VERTICAL_TOP);

/*
 * Example 1. Demonstrates a simple cell comments without formatting.
 */
/* Set up some worksheet formatting. */
$worksheet1->setColumn([2, 2], 25);
$worksheet1->setRow(2, 50);
$worksheet1->writeString("C3", "Hold the mouse over this cell to see the comment.", $text_wrap);
$worksheet1->writeComment("C3", "This is a comment.");

/*
 * Example 2. Demonstrates visible and hidden comments.
 */
/* Set up some worksheet formatting. */
$worksheet2->setColumn([2, 2], 25);
$worksheet2->setRow(2, 50);
$worksheet2->setRow(2, 50);
$worksheet2->writeString("C3", "This cell comment is visible.", $text_wrap);
/* Use an option to make the comment visible. */
$options2 = new CommentOptions();
$options2->visible = CommentDisplay::VISIBLE;
$worksheet2->writeComment("C3", "Hello.", $options2);
$worksheet2->writeString(
    "C6",
    "This cell comment isn't visible until you pass the" .
    " mouse over it (the default).",
    $text_wrap
);
$worksheet2->writeComment("C6", "Hello.");

/*
 * Example 3. Demonstrates visible and hidden comments, set at the
 * worksheet level.
 */
$worksheet3->setColumn([2, 2], 25);
$worksheet3->setRow(2, 50);
$worksheet3->setRow(5, 50);
$worksheet3->setRow(8, 50);
/* Make all comments on the worksheet visible. */
$worksheet3->showComments();
$worksheet3->writeString("C3", "This cell comment is visible, explicitly.", $text_wrap);

$options3a = new CommentOptions();
$options3a->visible = CommentDisplay::VISIBLE;

$worksheet3->writeComment([2, 2], "Hello", $options3a);
$worksheet3->writeString(
    "C6",
    "This cell comment is also visible because we used showComments().",
    $text_wrap
);
$worksheet3->writeComment("C6", "Hello");
$worksheet3->writeString("C9", "However, we can still override it locally.", $text_wrap);
$options3b = new CommentOptions();
$options3b->visible = CommentDisplay::HIDDEN;
$worksheet3->writeComment("C9", "Hello", $options3b);

/*
 * Example 4. Demonstrates changes to the comment box dimensions.
 */
$worksheet4->setColumn([2, 2], 25);
$worksheet4->setRow(2, 50);
$worksheet4->setRow(5, 50);
$worksheet4->setRow(8, 50);
$worksheet4->setRow(15, 50);
$worksheet4->setRow(18, 50);
$worksheet4->showComments();
$worksheet4->writeString("C3", "This cell comment is default size.", $text_wrap);
$worksheet4->writeComment([2, 2], "Hello");
$worksheet4->writeString("C6", "This cell comment is twice as wide.", $text_wrap);
$options4a = new CommentOptions();
$options4a->x_scale = 2.0;
$worksheet4->writeComment("C6", "Hello", $options4a);
$worksheet4->writeString("C9", "This cell comment is twice as high.", $text_wrap);
$options4b = new CommentOptions();
$options4b->y_scale = 2.0;
$worksheet4->writeComment("C9", "Hello", $options4b);
$worksheet4->writeString("C16", "This cell comment is scaled in both directions.", $text_wrap);
$options4c = new CommentOptions();
$options4c->x_scale = 1.2;
$options4c->y_scale = 0.5;
$worksheet4->writeComment("C16", "Hello", $options4c);
$worksheet4->writeString("C19", "This cell comment has width and height specified in pixels.", $text_wrap);
$options4d = new CommentOptions();
$options4d->width = 200;
$options4d->height = 50;
$worksheet4->writeComment("C19", "Hello", $options4d);

/*
 * Example 5. Demonstrates changes to the cell comment position.
 */
$worksheet5->setColumn([2, 2], 25);
$worksheet5->setRow(2, 50);
$worksheet5->setRow(5, 50);
$worksheet5->setRow(8, 50);
$worksheet5->showComments();
$worksheet5->writeString("C3", "This cell comment is in the default position.", $text_wrap);
$worksheet5->writeComment([2, 2], "Hello");
$worksheet5->writeString("C6", "This cell comment has been moved to another cell.", $text_wrap);
$options5a = new CommentOptions();
$options5a->start_row = 3;
$options5a->start_col = 4;
$worksheet5->writeComment("C6", "Hello", $options5a);
$worksheet5->writeString("C9", "This cell comment has been shifted within its default cell.", $text_wrap);
$options5b = new CommentOptions();
$options5b->start_row = 30;
$options5b->start_col = 12;
$worksheet5->writeComment("C9", "Hello", $options5b);

/*
 * Example 6. Demonstrates changes to the comment background color.
 */
$worksheet6->setColumn([2, 2], 25);
$worksheet6->setRow(2, 50);
$worksheet6->setRow(5, 50);
$worksheet6->setRow(8, 50);
$worksheet6->showComments();
$worksheet6->writeString("C3", "This cell comment has a different color.", $text_wrap);
$options6a = new CommentOptions();
$options6a->color = Color::GREEN;
$worksheet6->writeComment([2, 2], "Hello", $options6a);
$worksheet6->writeString("C6", "This cell comment has the default color.", $text_wrap);
$worksheet6->writeComment("C6", "Hello");
$worksheet6->writeString("C9", "This cell comment has a different color.", $text_wrap);
$options6b = new CommentOptions();
$options6b->color = 0xFF6600;
$worksheet6->writeComment("C9", "Hello", $options6b);

/*
 * Example 7. Demonstrates how to set the cell comment author.
 */
$worksheet7->setColumn([2, 2], 30);
$worksheet7->setRow(2, 50);
$worksheet7->setRow(5, 60);
$worksheet7->writeString("C3", "Move the mouse over this cell and you will" .
    " see 'Cell C3  commented by' (blank) " . "in the status bar at the bottom.", $text_wrap);
$worksheet7->writeComment("C3", "Hello");
$worksheet7->writeString("C6", "Move the mouse over this cell and you will see 'Cell C6" .
    " commented by libxlsxwriter' in the status " .
    "bar at the bottom.", $text_wrap);
$options7a = new CommentOptions();
$options7a->author = 'libxlsxwriter';
$worksheet7->writeComment("C6", "Hello", $options7a);

/*
 * Example 8. Demonstrates the need to explicitly set the row height.
 */
$worksheet8->setColumn([2, 2], 25);
$worksheet8->setRow(2, 80);
$worksheet8->showComments();
$worksheet8->writeString(
    "C3",
    "The height of this row has been adjusted explicitly" .
    " using worksheet_set_row(). The size of the comment box is" .
    " adjusted accordingly by libxlsxwriter",
    $text_wrap
);
$worksheet8->writeComment("C3", "Hello");
$worksheet8->writeString(
    "C6",
    "The height of this row has been adjusted by Excel when the " .
    "file is opened due to the text wrap property being set. " .
    "Unfortunately this means that the height of the row is " .
    "unknown to libxlsxwriter at run time and thus the comment " .
    "box is stretched as well.\n\n" .
    "Use worksheet_set_row() to specify the row height explicitly " .
    "to avoid this problem.",
    $text_wrap
);
$worksheet8->writeComment("C6", "Hello");

$workbook->close();
