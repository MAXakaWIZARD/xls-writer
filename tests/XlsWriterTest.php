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
        @unlink($this->testFilePath);
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
}
