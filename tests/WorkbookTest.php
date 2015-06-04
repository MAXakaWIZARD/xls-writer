<?php
namespace Test;

use Xls\Workbook;

/**
 *
 */
class WorkbookTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var Workbook
     */
    protected $workbook;

    /**
     * @var string
     */
    protected $testFilePath;

    /**
     *
     */
    public function setUp()
    {
        $this->testFilePath = TEST_DATA_PATH . '/test.xls';
        $this->workbook = new Workbook();
    }

    /**
     *
     */
    public function testDuplicateSheetName()
    {
        $sheetName = 'Sheet1';
        $this->workbook->addWorksheet($sheetName);
        $this->assertTrue($this->workbook->hasSheet($sheetName));

        $this->setExpectedException('\Exception', "Worksheet 'Sheet1' already exists");
        $this->workbook->addWorksheet($sheetName);
    }

    /**
     *
     */
    public function testActiveAndFirstSheet()
    {
        $firstSheet = $this->workbook->addWorksheet('Sheet1');
        $this->assertSame(0, $this->workbook->getActiveSheetIndex());
        $this->assertSame(0, $this->workbook->getFirstSheetIndex());
        $this->assertTrue($firstSheet->isSelected());

        $secondSheet = $this->workbook->addWorksheet('Sheet2');
        $secondSheet->activate();
        $this->assertSame(1, $this->workbook->getActiveSheetIndex());
        $this->assertTrue($secondSheet->isSelected());
        $this->assertFalse($firstSheet->isSelected());
        $secondSheet->setFirstSheet();
        $this->assertSame(1, $this->workbook->getFirstSheetIndex());

        $firstSheet->activate();
        $this->assertSame(0, $this->workbook->getActiveSheetIndex());
        $this->assertTrue($firstSheet->isSelected());
        $this->assertFalse($secondSheet->isSelected());
        $firstSheet->setFirstSheet();
        $this->assertSame(0, $this->workbook->getFirstSheetIndex());
    }

    /**
     *
     */
    public function testVeryLongSheetName()
    {
        $longName = str_repeat('a', 300);
        $this->setExpectedException('\Exception', "Sheet name must be shorter than 255 chars");
        $this->workbook->addWorksheet($longName);
    }

    /**
     *
     */
    public function testInvalidColorIndex()
    {
        $this->setExpectedException('\Exception', 'Color index 65 outside range: 8 <= index <= 64');
        $this->workbook->setCustomColor(65, 204, 204, 204);
    }

    /**
     *
     */
    public function testInvalidColor()
    {
        $this->setExpectedException('\Exception', 'Color component outside range: 0 <= color <= 255');
        $this->workbook->setCustomColor(12, 265, 265, 265);
    }

    /**
     *
     */
    public function testNoSheets()
    {
        $this->setExpectedException('\Exception', 'Cannot save workbook with no sheets');
        $this->workbook->save($this->testFilePath);
    }

    /**
     *
     */
    public function testWrongZoomFactor()
    {
        $sheet = $this->workbook->addWorksheet();
        $this->setExpectedException('\Exception', 'Zoom factor 1000 outside range: 10 <= zoom <= 400');
        $sheet->getPageSetup()->setZoom(1000);
    }

    /**
     *
     */
    public function testWrongPrintScaleFactor()
    {
        $sheet = $this->workbook->addWorksheet();
        $this->setExpectedException('\Exception', 'Print scale 1000 outside range: 10 <= zoom <= 400');
        $sheet->getPageSetup()->setPrintScale(1000);
    }

    /**
     *
     */
    public function testWriteRowWithWrongData()
    {
        $sheet = $this->workbook->addWorksheet();

        $this->setExpectedException('\Exception', '$val needs to be an array');
        $sheet->writeRow(0, 0, null);
    }

    /**
     *
     */
    public function testWriteColWithWrongData()
    {
        $sheet = $this->workbook->addWorksheet();

        $this->setExpectedException('\Exception', '$val needs to be an array');
        $sheet->writeCol(0, 0, null);
    }

    /**
     *
     */
    public function testWrongImagePath()
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
    public function testEmptyBmp()
    {
        $sheet = $this->workbook->addWorksheet();

        $path = TEST_DATA_PATH . '/corrupted.bmp';
        $this->setExpectedException('\Exception', "$path doesn't contain enough data");
        $sheet->insertBitmap(0, 0, $path);
    }

    /**
     *
     */
    public function testWrongRowIndex()
    {
        $sheet = $this->workbook->addWorksheet();

        $this->setExpectedException('\Exception', "Row index is beyond max row number");
        $sheet->write(100500, 0, 'Test');
    }

    /**
     *
     */
    public function testWrongColIndex()
    {
        $sheet = $this->workbook->addWorksheet();

        $this->setExpectedException('\Exception', "Col index is beyond max col number");
        $sheet->write(0, 100500, 'Test');
    }

    /**
     *
     */
    public function testBitmapInHiddenCell()
    {
        $sheet = $this->workbook->addWorksheet();

        $sheet->setRowHeight(0, 0);
        $this->setExpectedException('\Exception', "Bitmap isn't allowed to start or finish in a hidden cell");
        $sheet->insertBitmap(0, 0, TEST_DATA_PATH . '/elephpant.bmp');
    }

    /**
     * Valid formula should start with = or @
     */
    public function testInvalidFormula()
    {
        $sheet = $this->workbook->addWorksheet();

        $this->setExpectedException('\Exception', "Invalid formula: should start with = or @");
        $sheet->writeFormula(0, 0, 'A1+B1');
    }

    /**
     *
     */
    public function testInvalidTextRotation()
    {
        $format = $this->workbook->addFormat();

        $this->setExpectedException(
            '\Exception',
            "Invalid value for angle. Possible values are: 0, 90, 270 and -1 for stacking top-to-bottom."
        );
        $format->setTextRotation(30);
    }
}
