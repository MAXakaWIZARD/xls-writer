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
    protected function getFilePath($prefix, $suffix)
    {
        return TEST_DATA_PATH . '/' . $prefix . $suffix . '.xls';
    }

    /**
     * @param $params
     *
     * @return Workbook
     */
    protected function createWorkbook($params)
    {
        $workbook = new Workbook($params['format']);
        $workbook->setCreationTimestamp(self::WORKBOOK_TS);

        return $workbook;
    }

    /**
     * @return array
     */
    public function providerBiff5()
    {
        return array(
            array(
                array(
                    'format' => Biff5::VERSION,
                    'suffix' => ''
                )
            )
        );
    }

    /**
     * @return array
     */
    public function providerBiff8()
    {
        return array(
            array(
                array(
                    'format' => Biff8::VERSION,
                    'suffix' => '_biff8'
                )
            )
        );
    }

    /**
     * @return array
     */
    public function providerBiff5AndBiff8()
    {
        $biff5 = $this->providerBiff5();
        $biff8 = $this->providerBiff8();

        return array(
            $biff5[0],
            $biff8[0]
        );
    }
}
