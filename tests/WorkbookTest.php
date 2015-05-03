<?php
namespace Xls\Tests;

use Xls\Workbook;
use Xls\Biff5;
use Xls\Biff8;

/**
 *
 */
class WorkbookTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var string
     */
    protected $testFilePath;

    /**
     * @var Workbook
     */
    protected $workbook;

    /**
     *
     */
    public function setUp()
    {
        $this->testFilePath = TEST_DATA_PATH . '/test.xls';
        $this->workbook = new Workbook($this->testFilePath);
    }

    /**
     *
     */
    public function testDuplicateSheetName()
    {
        $sheetName = 'Sheet1';
        $this->workbook->addWorksheet($sheetName);
        $this->assertTrue($this->workbook->hasSheet($sheetName));

        $this->setExpectedException('\Exception', "Worksheet '$sheetName' already exists");
        $this->workbook->addWorksheet($sheetName);
    }

    /**
     *
     */
    public function testActiveAndFirstSheet()
    {
        $firstSheet = $this->workbook->addWorksheet('Sheet1');
        $this->assertSame(0, $this->workbook->getActiveSheet());
        $this->assertSame(0, $this->workbook->getFirstSheet());
        $this->assertTrue($firstSheet->isSelected());

        $secondSheet = $this->workbook->addWorksheet('Sheet2');
        $secondSheet->activate();
        $this->assertSame(1, $this->workbook->getActiveSheet());
        $this->assertTrue($secondSheet->isSelected());
        $this->assertFalse($firstSheet->isSelected());
        $secondSheet->setFirstSheet();
        $this->assertSame(1, $this->workbook->getFirstSheet());

        $firstSheet->activate();
        $this->assertSame(0, $this->workbook->getActiveSheet());
        $this->assertTrue($firstSheet->isSelected());
        $this->assertFalse($secondSheet->isSelected());
        $firstSheet->setFirstSheet();
        $this->assertSame(0, $this->workbook->getFirstSheet());
    }

    /**
     *
     */
    public function testLongSheetName()
    {
        $longName = str_repeat('a', 32);
        $this->setExpectedException('\Exception', "Sheet name must be shorter than 31 chars");
        $this->workbook->addWorksheet($longName);
    }

    /**
     *
     */
    public function testVeryLongSheetName()
    {
        $this->workbook = new Workbook($this->testFilePath, Biff8::VERSION);
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
}
