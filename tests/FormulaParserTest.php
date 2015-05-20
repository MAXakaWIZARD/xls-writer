<?php
namespace Test;

use Xls\Biff8;
use Xls\FormulaParser;
use Xls\Token;
use Xls\BIFFwriter;

/**
 *
 */
class FormulaParserTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @var FormulaParser
     */
    protected $parser;

    /**
     *
     */
    public function setUp()
    {
        $this->parser = new FormulaParser(BIFFwriter::BYTE_ORDER_LE);
    }

    /**
     * @dataProvider providerGeneral
     */
    public function testGeneral($params)
    {
        $this->parser->setExtSheet('Sheet1', 0);
        $this->parser->setExtSheet('Sheet2', 1);

        if (isset($params['correct']) && $params['correct'] === false) {
            $message = (isset($params['error'])) ? $params['error'] : '';
            $this->setExpectedException('\Exception', $message);
        }

        if (!is_array($params['formula'])) {
            $params['formula'] = array($params['formula']);
        }

        foreach ($params['formula'] as $formula) {
            $this->parser->getReversePolish($formula);
        }
    }

    /**
     * @return array
     */
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
                        '$F$2:$F$5',
                        'Sheet1!A1+Sheet1:Sheet2!B1',
                        "'Sheet1'!A1-'Sheet1:Sheet2'!A10",
                        '-2+2*3/4',
                        'IF(3>=2,1,0)',
                        'IF(3>2,1,0)',
                        'IF(3<2,1,0)',
                        'IF(3<=2,1,0)',
                        'IF(3=2,"Equal","Not equal")',
                        'IF(3<>2;1;0)',
                        '"Lazy dog " & "jumped over"',
                        'A3',
                        'SUM(Sheet2:Sheet1!A1:D4)',
                        '0',
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
            ),
            array(
                array(
                    'formula' => '(2+3',
                    'correct' => false,
                    'error' => "')' token expected."
                )
            ),
            array(
                array(
                    'formula' => 'LEN()',
                    'correct' => false,
                    'error' => 'Incorrect number of arguments in function LEN()'
                )
            ),
            array(
                array(
                    'formula' => 'WHATEVERFUNCTION()',
                    'correct' => false,
                    'error' => "Function WHATEVERFUNCTION() doesn't exist"
                )
            ),
            array(
                array(
                    'formula' => 'Sheet3!A1+Sheet1:Sheet2!B1',
                    'correct' => false,
                    'error' => "Unknown sheet name Sheet3 in formula"
                )
            ),
            array(
                array(
                    'formula' => 'Sheet1!A1+Sheet3:Sheet1!B1',
                    'correct' => false,
                    'error' => "Unknown sheet name Sheet3 in formula"
                )
            ),
            array(
                array(
                    'formula' => 'Sheet1!A1+Sheet1:Sheet4!B1',
                    'correct' => false,
                    'error' => "Unknown sheet name Sheet4 in formula"
                )
            ),
            array(
                array(
                    'formula' => 'A100500+B100500',
                    'correct' => false,
                    'error' => "Row in: A100500 greater than " . (Biff8::MAX_ROWS - 1)
                )
            ),
            array(
                array(
                    'formula' => 'ZZ1+A1',
                    'correct' => false,
                    'error' => "Column in: ZZ1 greater than " . (Biff8::MAX_COLS - 1)
                )
            ),
            array(
                array(
                    'formula' => '"' . str_repeat('a', 300) . '" & ""',
                    'correct' => false,
                    'error' => "String is too long"
                )
            ),
        );
    }

    /**
     *
     */
    public function testToken()
    {
        $this->assertEquals(true, Token::isRangeWithColon('F2:F5'), 'F2:F5 is valid range');
        $this->assertEquals(true, Token::isRangeWithColon('$F$2:$F$5'), '$F$2:$F$5 is valid range');

        $this->assertEquals(true, Token::isRangeWithDots('F2..F5'), 'F2..F5 is valid range');
        $this->assertEquals(true, Token::isRangeWithDots('$F$2..$F$5'), '$F$2..$F$5 is valid range');
    }
}
