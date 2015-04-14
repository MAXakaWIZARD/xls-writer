<?php
namespace Xls\Writer\Tests;

use Xls\Writer;

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

        $worksheet = $workbook->addWorksheet('My first worksheet');

        $formatHeader = $workbook->addFormat();
        $formatHeader->setBold();

        $worksheet->write(0, 0, 'Title', $formatHeader);
        $worksheet->write(0, 1, 'Count', $formatHeader);
        $worksheet->write(0, 2, 'Price', $formatHeader);
        $worksheet->write(0, 3, 'Amount', $formatHeader);
        $worksheet->write(1, 0, 'Intel Core i7 2600K');
        $worksheet->write(1, 1, 1);
        $worksheet->write(1, 2, 500);
        $worksheet->writeFormula(1, 3, '=B2*C2');
        $worksheet->write(2, 0, 'ASUS P8P67');
        $worksheet->write(2, 1, 1);
        $worksheet->write(2, 2, 325);
        $worksheet->writeFormula(2, 3, '=B3*C3');
        $worksheet->write(3, 0, 'DDR2-800 8Gb');
        $worksheet->write(3, 1, 4);
        $worksheet->write(3, 2, 100);
        $worksheet->writeFormula(3, 3, '=B4*C4');

        $worksheet->write(10, 0, 'Total');
        $worksheet->writeFormula(10, 3, '=SUM(D2:D9)');

        // We still need to explicitly close the workbook
        $workbook->close();

        $this->assertFileExists($this->testFilePath);
        $this->assertFileEquals(TEST_DATA_PATH . '/formulas.xls', $this->testFilePath);
    }
}
