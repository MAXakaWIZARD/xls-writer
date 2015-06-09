<?php
namespace Test;

/**
 *
 */
class BitmapTest extends TestAbstract
{
    /**
     *
     */
    public function testWrongPath()
    {
        $sheet = $this->workbook->addWorksheet();

        $path = TEST_DATA_PATH . '/not_exists.bmp';
        $this->setExpectedException('\Exception', "Couldn't import $path");
        $sheet->insertBitmap(0, 0, $path);
    }

    /**
     *
     */
    public function testNonBmp()
    {
        $sheet = $this->workbook->addWorksheet();

        $path = TEST_DATA_PATH . '/fill.xls';
        $this->setExpectedException('\Exception', "$path doesn't appear to be a valid bitmap image");
        $sheet->insertBitmap(0, 0, $path);
    }

    /**
     *
     */
    public function testEmpty()
    {
        $sheet = $this->workbook->addWorksheet();

        $path = TEST_DATA_PATH . '/corrupted.bmp';
        $this->setExpectedException('\Exception', "$path doesn't contain enough data");
        $sheet->insertBitmap(0, 0, $path);
    }

    /**
     *
     */
    public function test16bit()
    {
        $sheet = $this->workbook->addWorksheet();

        $path = TEST_DATA_PATH . '/elephpant_16bit.bmp';
        $this->setExpectedException('\Exception', "$path isn't a 24bit true color bitmap");
        $sheet->insertBitmap(0, 0, $path);
    }

    /**
     *
     */
    public function test2planes()
    {
        $sheet = $this->workbook->addWorksheet();

        $path = TEST_DATA_PATH . '/elephpant_2planes.bmp';
        $this->setExpectedException('\Exception', "$path: only 1 plane supported in bitmap image");
        $sheet->insertBitmap(0, 0, $path);
    }

    /**
     *
     */
    public function testCompressed()
    {
        $sheet = $this->workbook->addWorksheet();

        $path = TEST_DATA_PATH . '/elephpant_compressed.bmp';
        $this->setExpectedException('\Exception', "$path: compression not supported in bitmap image");
        $sheet->insertBitmap(0, 0, $path);
    }

    /**
     *
     */
    public function testHiddenCell()
    {
        $sheet = $this->workbook->addWorksheet();

        $sheet->setRowHeight(0, 0);
        $this->setExpectedException('\Exception', "Bitmap isn't allowed to start or finish in a hidden cell");
        $sheet->insertBitmap(0, 0, TEST_DATA_PATH . '/elephpant.bmp');
    }

    public function testImage()
    {
        $sheet = $this->workbook->addWorksheet();
        $sheet->write(0, 0, 'Test');
        $sheet->insertBitmap(2, 2, TEST_DATA_PATH . '/elephpant.bmp');

        $this->workbook->save($this->testFilePath);

        $this->assertTestFileEqualsTo('image');
    }
}
