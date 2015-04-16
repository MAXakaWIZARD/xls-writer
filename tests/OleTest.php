<?php
namespace Xls\Writer\Tests;

use Xls\Writer;
use Xls\OLE;
use Xls\Writer\Format;

/**
 *
 */
class OleTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var OLE
     */
    protected $ole;

    /**
     *
     */
    public function setUp()
    {
        $this->ole = new OLE();
    }

    /**
     *
     */
    public function tearDown()
    {

    }

    /**
     *
     */
    public function testNonExistentFile()
    {
        $filePath = TEST_DATA_PATH . '/not_exists.xls';

        $this->setExpectedException('\Exception', 'Can\'t open file ' . $filePath);
        $this->ole->read($filePath);
    }

    /**
     *
     */
    public function testNonOle()
    {
        $filePath = TEST_DATA_PATH . '/non_ole.xls';

        $this->setExpectedException('\Exception', "File doesn't seem to be an OLE container.");
        $this->ole->read($filePath);
    }

    /**
     *
     */
    public function testNonLittleEndian()
    {
        $filePath = TEST_DATA_PATH . '/non_little_endian.xls';

        $this->setExpectedException('\Exception', "Only Little-Endian encoding is supported.");
        $this->ole->read($filePath);
    }

    /**
     *
     */
    public function testGeneral()
    {
        //$filePath = TEST_DATA_PATH . '/general.xls';

        //$this->ole->read($filePath);
    }
}
