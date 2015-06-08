<?php
namespace Test;

use Xls\Cell;
use Xls\Utils;
use Xls\Token;

/**
 *
 */
class UtilsTest extends TestAbstract
{
    /**
     * @dataProvider providerRowColToCell
     * @param $expected
     * @param $row
     * @param $col
     *
     * @throws \Exception
     */
    public function testRowColToCell($expected, $row, $col)
    {
        $cell = new Cell($row, $col);
        $this->assertSame($expected, $cell->getAddress());
    }

    /**
     * @return array
     */
    public function providerRowColToCell()
    {
        return array(
            array('A1', 0, 0),
            array('B2', 1, 1),
            array('K256', 255, 10),
            array('IV256', 255, 255),
            array('AB1', 0, 27),
        );
    }

    /**
     * @dataProvider providerAddressToRowCol
     *
     * @param string $address
     * @param array $expected
     */
    public function testAddressToRowCol($address, $expected)
    {
        $this->assertEquals($expected, Cell::addressToRowCol($address), 'Wrong value for cell address: ' . $address);
    }

    /**
     * @return array
     */
    public function providerAddressToRowCol()
    {
        return array(
            array('A1', array(0, 0, 1, 1)),
            array('B2', array(1, 1, 1, 1)),
            array('$B2', array(1, 1, 1, 0)),
            array('C$3', array(2, 2, 0, 1)),
            array('$C$3', array(2, 2, 0, 0)),
            array('K256', array(255, 10, 1, 1)),
            array('IV256', array(255, 255, 1, 1)),
            array('AB1', array(0, 27, 1, 1)),
            array('ZZ257', array(256, 701, 1, 1)),
            array('ZZZ257', array(256, 18277, 1, 1)),
        );
    }

    /**
     * @dataProvider providerHexDump
     * @param string $expected
     * @param int $value
     *
     * @throws \Exception
     */
    public function testHexDump($expected, $value)
    {
        $this->assertSame($expected, Utils::hexDump($value));
    }

    /**
     * @return array
     */
    public function providerHexDump()
    {
        return array(
            array('FF 00', pack('v', 255)),
            array('01 00 00 00', pack('V', 1)),
        );
    }

    public function testGuids()
    {
        $guid = Utils::generateGuid();
        $this->assertEquals(32, strlen($guid), 'GUID mus be 32 chars long');

        $this->assertNotSame($guid, Utils::generateGuid(), 'Every new GUID must be unique');
    }
}
