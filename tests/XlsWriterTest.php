<?php
namespace Xls\Writer\Tests;

use Xls\Writer;
use Xls\Writer\Format;

/**
 *
 */
class XlsWriterTest extends \PHPUnit_Framework_TestCase
{
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
    }

    /**
     *
     */
    public function tearDown()
    {
        //@unlink($this->testFilePath);
    }

    /**
     *
     */
    public function testGeneral()
    {
        $workbook = new \Xls\Writer($this->testFilePath);

        //needed for test files comparison
        $workbook->setCreationTimestamp(1429042916);

        $worksheet = $workbook->addWorksheet('My first worksheet');

        $worksheet->write(0, 0, 'Name');
        $worksheet->write(0, 1, 'Age');
        $worksheet->write(1, 0, 'John Smith');
        $worksheet->write(1, 1, 30);
        $worksheet->write(2, 0, 'Johann Schmidt');
        $worksheet->write(2, 1, 31);
        $worksheet->write(3, 0, 'Juan Herrera');
        $worksheet->write(3, 1, 32);

        // We still need to explicitly close the workbook
        $workbook->close();

        $this->assertFileExists($this->testFilePath);
        $this->assertFileEquals(TEST_DATA_PATH . '/general.xls', $this->testFilePath);
    }

    /**
     *
     */
    public function testRich()
    {
        $workbook = new \Xls\Writer($this->testFilePath);

        //needed for test files comparison
        $workbook->setCreationTimestamp(1429042916);

        $worksheet = $workbook->addWorksheet('New PC');

        $headerFormat = $workbook->addFormat();
        $headerFormat->setBold();
        $headerFormat->setBorder(Format::BORDER_THIN);
        $headerFormat->setColor('blue');
        $headerFormat->setFgColor('silver');

        $cellFormat = $workbook->addFormat();
        $cellFormat->setBorder(Format::BORDER_THIN);

        $worksheet->write(0, 0, 'Title', $headerFormat);
        $worksheet->write(0, 1, 'Count', $headerFormat);
        $worksheet->write(0, 2, 'Price', $headerFormat);
        $worksheet->write(0, 3, 'Amount', $headerFormat);
        $worksheet->write(1, 0, 'Intel Core i7 2600K', $cellFormat);
        $worksheet->write(1, 1, 1, $cellFormat);
        $worksheet->write(1, 2, 500, $cellFormat);
        $worksheet->writeFormula(1, 3, '=B2*C2', $cellFormat);
        $worksheet->write(2, 0, 'ASUS P8P67', $cellFormat);
        $worksheet->write(2, 1, 1, $cellFormat);
        $worksheet->write(2, 2, 325, $cellFormat);
        $worksheet->writeFormula(2, 3, '=B3*C3', $cellFormat);
        $worksheet->write(3, 0, 'DDR2-800 8Gb', $cellFormat);
        $worksheet->write(3, 1, 4, $cellFormat);
        $worksheet->write(3, 2, 100, $cellFormat);
        $worksheet->writeFormula(3, 3, '=B4*C4', $cellFormat);

        $totalFormat = $workbook->addFormat();
        $totalFormat->setBold();
        $totalFormat->setBorder(Format::BORDER_THIN);

        for ($i = 4; $i < 10; $i++) {
            for ($j = 0; $j < 4; $j++) {
                $worksheet->write($i, $j, '', $cellFormat);
            }
        }

        $worksheet->write(10, 0, 'Total', $totalFormat);
        $worksheet->write(10, 1, '', $totalFormat);
        $worksheet->write(10, 2, '', $totalFormat);
        $worksheet->mergeCells(10, 0, 10, 2);
        $worksheet->writeFormula(10, 3, '=SUM(D2:D9)', $totalFormat);

        // We still need to explicitly close the workbook
        $workbook->close();

        $this->assertFileExists($this->testFilePath);
        $this->assertFileEquals(TEST_DATA_PATH . '/rich.xls', $this->testFilePath);
    }
}
