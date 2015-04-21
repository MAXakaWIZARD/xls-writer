<?php
namespace Xls\Writer\Tests;

use Xls\Writer\Parser;
use Xls\Writer\BIFFwriter;

/**
 *
 */
class FormulaParserTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var Parser
     */
    protected $parser;

    /**
     *
     */
    public function setUp()
    {
        $this->parser = new Parser(BIFFwriter::BYTE_ORDER_LE, BIFFwriter::VERSION_5);
    }

    /**
     * @dataProvider providerGeneral
     */
    public function testGeneral($params)
    {
        if (isset($params['correct']) && $params['correct'] === false) {
            $message = (isset($params['error'])) ? $params['error'] : '';
            $this->setExpectedException('\Exception', $message);
        }

        if (!is_array($params['formula'])) {
            $params['formula'] = array($params['formula']);
        }

        foreach ($params['formula'] as $formula) {
            $result = $this->parser->parse($formula);

            if (isset($params['correct']) && $params['correct'] === true) {
                $this->assertTrue($result);
            }
        }
    }

    public function providerGeneral()
    {
        return array(
            array(
                array(
                    'formula' => array(
                        'SUM(D2:D9)',
                        'SUM(D2..D9)',
                        '(A1+A2)*(B1-B5)',
                        'C3/C4',
                        '$C$2+$D$3',
                        'Sheet1!A1+Sheet1:Sheet2!B1',
                        "'Sheet1'!A1-'Sheet1:Sheet2'!A10",
                        '2+2*3/4'
                    ),
                    'correct' => true
                )
            ),
            array(
                array(
                    'formula' => '=SUM(D2:D9)',
                    'correct' => false,
                    'error' => 'Syntax error: =, lookahead: S, current char: 1'
                )
            ),
            array(
                array(
                    'formula' => 'SUM(D2:D9',
                    'correct' => false,
                    'error' => 'Syntax error: comma expected in function SUM, arg #1'
                )
            ),
            array(
                array(
                    'formula' => '2**3',
                    'correct' => false,
                    'error' => 'Syntax error: *, lookahead: 3, current char: 3'
                )
            )
        );
    }
}