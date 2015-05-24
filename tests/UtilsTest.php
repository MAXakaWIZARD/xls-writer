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
     *
     */
    public function testRowColToCellInvalid()
    {
        $this->setExpectedException('\Exception', 'Maximum column value exceeded: 256');
        Cell::getAddress(0, 256);
    }

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
        $this->assertSame($expected, Cell::getAddress($row, $col));
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

    /**
     * @dataProvider providerTokenGetPtg
     * @param string $expected
     * @param string $value
     *
     * @throws \Exception
     */
    public function testTokenGetPtg($expected, $value)
    {
        $this->assertSame($expected, Token::getPtg($value));
    }

    /**
     * @return array
     */
    public function providerTokenGetPtg()
    {
        return array(
            array(null, 'UNKNOWN_TOKEN'),
            array('ptgAdd', Token::TOKEN_ADD),
        );
    }

    /**
     * @dataProvider providerTokenPossibleLookahead
     * @param string $expected
     * @param string $token
     * @param string $lookahead
     *
     * @throws \Exception
     */
    public function testTokenPossibleLookahead($expected, $token, $lookahead)
    {
        $this->assertSame($expected, Token::isPossibleLookahead($token, $lookahead));
    }

    /**
     * @return array
     */
    public function providerTokenPossibleLookahead()
    {
        return array(
            array(true, Token::TOKEN_GT, '='),
            array(false, Token::TOKEN_GT, '<'),
            array(true, Token::TOKEN_MUL, '5'),
        );
    }
}
