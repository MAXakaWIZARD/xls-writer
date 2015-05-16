<?php
namespace Test;

use Xls\Workbook;
use Xls\Biff5;
use Xls\Biff8;
use Xls\Format;
use Xls\Fill;
use Xls\Cell;

/**
 *
 */
class TestAbstract extends \PHPUnit_Framework_TestCase
{
    const WORKBOOK_TS = 1429042916;

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
     * @param $prefix
     * @param $suffix
     *
     * @return string
     */
    protected function getFilePath($prefix, $suffix = '')
    {
        return TEST_DATA_PATH . '/' . $prefix . $suffix . '.xls';
    }

    /**
     * @return Workbook
     */
    protected function createWorkbook()
    {
        $workbook = new Workbook();
        $workbook->setCreationTimestamp(self::WORKBOOK_TS);

        return $workbook;
    }

    /**
     * @return Workbook
     */
    protected function createWorkbookBiff5()
    {
        $workbook = new Workbook(Biff5::VERSION);
        $workbook->setCreationTimestamp(self::WORKBOOK_TS);

        return $workbook;
    }

    /**
     * @param string $name
     * @param string $suffix
     */
    protected function checkTestFileIsEqualTo($name, $suffix = '')
    {
        $this->assertFileExists($this->testFilePath);
        $correctFilePath = $this->getFilePath($name, $suffix);
        $this->assertFileEquals($correctFilePath, $this->testFilePath);
    }
}
