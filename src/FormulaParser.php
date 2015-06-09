<?php

namespace Xls;

class FormulaParser
{
    /**
     * The index of the character we are currently looking at
     * @var integer
     */
    protected $currentChar = 0;

    /**
     * The token we are working on.
     * @var string
     */
    protected $currentToken = '';

    /**
     * The formula to parse
     * @var string
     */
    protected $formula = '';

    /**
     * The character ahead of the current char
     * @var string
     */
    protected $lookahead = '';

    /**
     * The parse tree to be generated
     * @var array
     */
    protected $parseTree = array();

    /**
     * Array of external sheets
     * @var array
     */
    protected $extSheets = array();

    /**
     * Array of sheet references in the form of REF structures
     * @var array
     */
    protected $references = array();

    /**
     * Convert a token to the proper ptg value.
     *
     * @param mixed $token The token to convert.
     * @return string the converted token
     * @throws \Exception
     */
    protected function convert($token)
    {
        if (Token::isString($token)) {
            return $this->convertString($token);
        } elseif (is_numeric($token)) {
            return $this->convertNumber($token);
        } elseif (Token::isReference($token)) {
            return $this->convertRef2d($token);
        } elseif (Token::isExternalReference($token)) {
            return $this->convertRef3d($token);
        } elseif (Token::isRange($token)) {
            return $this->convertRange2d($token);
        } elseif (Token::isExternalRange($token)) {
            return $this->convertRange3d($token);
        } elseif (Ptg::exists($token)) {
            // operators (including parentheses)
            return pack("C", Ptg::get($token));
        } elseif (Token::isArg($token)) {
            // if it's an argument, ignore the token (the argument remains)
            return '';
        }

        throw new \Exception("Unknown token $token");
    }

    /**
     * Convert a number token to ptgInt or ptgNum
     *
     * @param mixed $num an integer or double for conversion to its ptg value
     * @return string
     */
    protected function convertNumber($num)
    {
        if (preg_match("/^\d+$/", $num) && $num <= 65535) {
            // Integer in the range 0..2**16-1
            return pack("Cv", Ptg::get('ptgInt'), $num);
        } else {
            // A float
            return pack("Cd", Ptg::get('ptgNum'), $num);
        }
    }

    /**
     * Convert a string token to ptgStr
     *
     * @param string $string A string for conversion to its ptg value.
     * @throws \Exception
     * @return string the converted token
     */
    protected function convertString($string)
    {
        // chop away beggining and ending quotes
        $string = substr($string, 1, strlen($string) - 2);
        if (strlen($string) > Biff8::MAX_STR_LENGTH) {
            throw new \Exception("String is too long");
        }

        $encoding = 0;

        return pack("CCC", Ptg::get('ptgStr'), strlen($string), $encoding) . $string;
    }

    /**
     * Convert a function to a ptgFunc or ptgFuncVarV depending on the number of
     * args that it takes.
     *
     * @param string $token    The name of the function for convertion to ptg value.
     * @param integer $numArgs The number of arguments the function receives.
     *
     * @return string The packed ptg for the function
     */
    protected function convertFunction($token, $numArgs)
    {
        $ptg = Functions::getPtg($token);
        $args = Functions::getArgsNumber($token);

        // Fixed number of args eg. TIME($i,$j,$k).
        if ($args >= 0) {
            return pack("Cv", Ptg::get('ptgFuncV'), $ptg);
        }

        // Variable number of args eg. SUM($i,$j,$k, ..).
        return pack("CCv", Ptg::get('ptgFuncVarV'), $numArgs, $ptg);
    }

    /**
     * Convert an Excel range such as A1:D4 to a ptgRefV.
     *
     * @param string $range An Excel range in the A1:A2 or A1..A2 format.
     * @return string
     */
    protected function convertRange2d($range)
    {
        $separator = (Token::isRangeWithDots($range)) ? '..' : ':';
        list($cell1, $cell2) = explode($separator, $range);

        // Convert the cell references
        list($row1, $col1) = $this->cellToPackedRowcol($cell1);
        list($row2, $col2) = $this->cellToPackedRowcol($cell2);

        $ptgArea = pack("C", Ptg::get('ptgArea'));

        return $ptgArea . $row1 . $row2 . $col1 . $col2;
    }

    /**
     * Convert an Excel 3d range such as "Sheet1!A1:D4" or "Sheet1:Sheet2!A1:D4" to
     * a ptgArea3d.
     *
     * @param string $token An Excel range in the Sheet1!A1:A2 format.
     * @return string The packed ptgArea3d token
     */
    protected function convertRange3d($token)
    {
        // Split the ref at the ! symbol
        list($extRef, $range) = explode('!', $token);

        // Convert the external reference part
        $extRef = $this->getRefIndex($extRef);

        // Split the range into 2 cell refs
        list($cell1, $cell2) = explode(':', $range);

        // Convert the cell references
        list($row1, $col1) = $this->cellToPackedRowcol($cell1);
        list($row2, $col2) = $this->cellToPackedRowcol($cell2);

        $ptgArea = pack("C", Ptg::get('ptgArea3dA'));

        return $ptgArea . $extRef . $row1 . $row2 . $col1 . $col2;
    }

    /**
     * Convert an Excel reference such as A1, $B2, C$3 or $D$4 to a ptgRefV.
     *
     * @param string $cell An Excel cell reference
     * @return string The cell in packed() format with the corresponding ptg
     */
    protected function convertRef2d($cell)
    {
        list($row, $col) = $this->cellToPackedRowcol($cell);

        $ptgRef = pack("C", Ptg::get('ptgRefA'));

        return $ptgRef . $row . $col;
    }

    /**
     * Convert an Excel 3d reference such as "Sheet1!A1" or "Sheet1:Sheet2!A1" to a
     * ptgRef3d.
     *
     * @param string $cell An Excel cell reference
     * @return string The packed ptgRef3d token
     */
    protected function convertRef3d($cell)
    {
        // Split the ref at the ! symbol
        list($extRef, $cell) = explode('!', $cell);

        // Convert the external reference part
        $extRef = $this->getRefIndex($extRef);

        // Convert the cell reference part
        list($row, $col) = $this->cellToPackedRowcol($cell);

        $ptgRef = pack("C", Ptg::get('ptgRef3dA'));

        return $ptgRef . $extRef . $row . $col;
    }

    /**
     * @param string $str
     *
     * @return string
     */
    protected function removeTrailingQuotes($str)
    {
        $str = preg_replace("/^'/", '', $str); // Remove leading  ' if any.
        $str = preg_replace("/'$/", '', $str); // Remove trailing ' if any.

        return $str;
    }

    /**
     * @param $extRef
     *
     * @return array
     * @throws \Exception
     */
    protected function getRangeSheets($extRef)
    {
        $extRef = $this->removeTrailingQuotes($extRef);

        // Check if there is a sheet range eg., Sheet1:Sheet2.
        if (preg_match("/:/", $extRef)) {
            list($sheetName1, $sheetName2) = explode(':', $extRef);

            $sheet1 = $this->getSheetIndex($sheetName1);
            if ($sheet1 == -1) {
                throw new \Exception("Unknown sheet name $sheetName1 in formula");
            }

            $sheet2 = $this->getSheetIndex($sheetName2);
            if ($sheet2 == -1) {
                throw new \Exception("Unknown sheet name $sheetName2 in formula");
            }

            // Reverse max and min sheet numbers if necessary
            if ($sheet1 > $sheet2) {
                list($sheet1, $sheet2) = array($sheet2, $sheet1);
            }
        } else { // Single sheet name only.
            $sheet1 = $this->getSheetIndex($extRef);
            if ($sheet1 == -1) {
                throw new \Exception("Unknown sheet name $extRef in formula");
            }
            $sheet2 = $sheet1;
        }

        return array($sheet1, $sheet2);
    }

    /**
     * Look up the REF index that corresponds to an external sheet name
     * (or range). If it doesn't exist yet add it to the workbook's references
     * array. It assumes all sheet names given must exist.
     *
     * @param string $extRef The name of the external reference
     *
     * @throws \Exception
     * @return string The reference index in packed() format
     */
    protected function getRefIndex($extRef)
    {
        list($sheet1, $sheet2) = $this->getRangeSheets($extRef);

        $index = $this->addRef($sheet1, $sheet2);

        return pack('v', $index);
    }

    /**
     * Add reference and return its index
     * @param int $sheet1
     * @param int $sheet2
     *
     * @return int
     */
    public function addRef($sheet1, $sheet2)
    {
        // assume all references belong to this document
        $supbookIndex = 0x00;
        $ref = pack('vvv', $supbookIndex, $sheet1, $sheet2);

        $index = array_search($ref, $this->references);
        if ($index === false) {
            // if REF was not found add it to references array
            $this->references[] = $ref;
            $index = count($this->references);
        }

        return $index;
    }

    /**
     * Look up the index that corresponds to an external sheet name. The hash of
     * sheet names is updated by the addworksheet() method of the
     * Workbook class.
     *
     * @param string $sheetName
     *
     * @return integer The sheet index, -1 if the sheet was not found
     */
    protected function getSheetIndex($sheetName)
    {
        if (!isset($this->extSheets[$sheetName])) {
            return -1;
        }

        return $this->extSheets[$sheetName];
    }

    /**
     * This method is used to update the array of sheet names. It is
     * called by the addWorksheet() method of the
     * Workbook class.
     *
     * @see Workbook::addWorksheet()
     * @param string $name  The name of the worksheet being added
     * @param integer $index The index of the worksheet being added
     */
    public function addSheet($name, $index)
    {
        $this->extSheets[$name] = $index;
    }

    /**
     * pack() row and column into the required 3 or 4 byte format.
     *
     * @param string $cellAddress The Excel cell reference to be packed
     *
     * @throws \Exception
     * @return array Array containing the row and column in packed() format
     */
    protected function cellToPackedRowcol($cellAddress)
    {
        $cellAddress = strtoupper($cellAddress);
        $cell = Cell::createFromAddress($cellAddress);

        // Set the high bits to indicate if row or col are relative.
        $col = $cell->getCol();
        $col |= (int)$cell->isColRelative() << 14;
        $col |= (int)$cell->isRowRelative() << 15;
        $col = pack('v', $col);

        $row = pack('v', $cell->getRow());

        return array($row, $col);
    }

    /**
     * Advance to the next valid token.
     *
     */
    protected function advance()
    {
        $token = '';

        $position = $this->eatWhitespace();
        $formulaLength = strlen($this->formula);

        while ($position < $formulaLength) {
            $token .= $this->formula[$position];
            if ($position < $formulaLength - 1) {
                $this->lookahead = $this->formula[$position + 1];
            } else {
                $this->lookahead = '';
            }

            if ($this->match($token) != '') {
                $this->currentChar = $position + 1;
                $this->currentToken = $token;
                return;
            }

            if ($position < ($formulaLength - 2)) {
                $this->lookahead = $this->formula[$position + 2];
            } else {
                // if we run out of characters lookahead becomes empty
                $this->lookahead = '';
            }
            $position++;
        }
    }

    /**
     * @return int
     */
    protected function eatWhitespace()
    {
        $position = $this->currentChar;
        $formulaLength = strlen($this->formula);

        // eat up white spaces
        if ($position < $formulaLength) {
            while ($this->formula{$position} == " ") {
                $position++;
            }

            if ($position < ($formulaLength - 1)) {
                $this->lookahead = $this->formula{$position + 1};
            }
        }

        return $position;
    }

    /**
     * Checks if it's a valid token.
     *
     * @param string $token The token to check.
     * @return string The checked token
     */
    protected function match($token)
    {
        if (Token::isDeterministic($token)) {
            return $token;
        }

        if (Token::isLtOrGt($token)) {
            if (!Token::isPossibleLookahead($token, $this->lookahead)) {
                // it's not a GE, LTE or NE token
                return $token;
            }

            return '';
        }

        return $this->processDefaultCase($token);
    }

    /**
     * @param string $token
     *
     * @return string
     */
    protected function processDefaultCase($token)
    {
        if ($this->isReference($token)
            || $this->isExternalReference($token)
            || $this->isAnyRange($token)
            || $this->isNumber($token)
            || Token::isString($token)
            || $this->isFunctionCall($token)
        ) {
            return $token;
        }

        return '';
    }

    /**
     * @return bool
     */
    protected function lookaheadHasNumber()
    {
        return preg_match("/[0-9]/", $this->lookahead) === 1;
    }

    /**
     * @return bool
     */
    protected function isLookaheadDotOrColon()
    {
        return $this->lookahead === '.' || $this->lookahead === ':';
    }

    /**
     * @param string $token
     *
     * @return bool
     */
    protected function isAnyRange($token)
    {
        return Token::isAnyRange($token)
            && !$this->lookaheadHasNumber();
    }

    /**
     * @param string $token
     *
     * @return bool
     */
    protected function isReference($token)
    {
        return Token::isReference($token)
            && !$this->lookaheadHasNumber()
            && !$this->isLookaheadDotOrColon()
            && $this->lookahead != '!';
    }

    /**
     * @param string $token
     *
     * @return bool
     */
    protected function isExternalReference($token)
    {
        return Token::isExternalReference($token)
            && !$this->lookaheadHasNumber()
            && !$this->isLookaheadDotOrColon();
    }

    /**
     * If it's a number (check that it's not a sheet name or range)
     * @param string $token
     *
     * @return bool
     */
    protected function isNumber($token)
    {
        return is_numeric($token)
            && (!is_numeric($token . $this->lookahead) || $this->lookahead == '')
            && $this->lookahead != '!'
            && $this->lookahead != ':';
    }

    /**
     * @param string $token
     *
     * @return bool
     */
    protected function isFunctionCall($token)
    {
        return Token::isFunctionCall($token)
            && $this->lookahead == "(";
    }

    /**
     * The parsing method. It parses a formula.
     *
     * @param string $formula The formula to parse, without the initial equal sign (=).
     */
    public function parse($formula)
    {
        $this->parseTree = array();
        $this->currentChar = 0;
        $this->currentToken = '';
        $this->formula = $formula;
        $this->lookahead = (isset($formula[1])) ? $formula[1] : '';
        $this->advance();
        $this->parseTree = $this->condition();
    }

    /**
     * It parses a condition. It assumes the following rule:
     * Cond -> Expr [(">" | "<") Expr]
     *
     * @return array The parsed ptg'd tree
     */
    protected function condition()
    {
        $result = $this->expression();

        if (Token::isComparison($this->currentToken) || Token::isConcat($this->currentToken)) {
            $ptg = Token::getPtg($this->currentToken);
            $this->advance();
            $result = $this->createTree($ptg, $result, $this->expression());
        }

        return $result;
    }

    /**
     * It parses a expression. It assumes the following rule:
     * Expr -> Term [("+" | "-") Term]
     *      -> "string"
     *      -> "-" Term
     *
     * @return array The parsed ptg'd tree
     */
    protected function expression()
    {
        // If it's a string return a string node
        if (Token::isString($this->currentToken)) {
            $result = $this->createTree($this->currentToken, '', '');
            $this->advance();

            return $result;
        } elseif ($this->currentToken == Token::TOKEN_SUB) {
            // catch "-" Term
            $this->advance();

            return $this->createTree('ptgUminus', $this->expression(), '');
        }

        $result = $this->term();

        while (Token::isAddOrSub($this->currentToken)) {
            $ptg = Token::getPtg($this->currentToken);
            $this->advance();
            $result = $this->createTree($ptg, $result, $this->term());
        }

        return $result;
    }

    /**
     * This function just introduces a ptgParen element in the tree, so that Excel
     * doesn't get confused when working with a parenthesized formula afterwards.
     *
     * @see _fact()
     * @return array The parsed ptg'd tree
     */
    protected function parenthesizedExpression()
    {
        return $this->createTree('ptgParen', $this->expression(), '');
    }

    /**
     * It parses a term. It assumes the following rule:
     * Term -> Fact [("*" | "/") Fact]
     *
     * @return array The parsed ptg'd tree
     */
    protected function term()
    {
        $result = $this->fact();

        while (Token::isMulOrDiv($this->currentToken)) {
            $ptg = Token::getPtg($this->currentToken);
            $this->advance();
            $result = $this->createTree($ptg, $result, $this->fact());
        }

        return $result;
    }

    /**
     * It parses a factor. It assumes the following rule:
     * Fact -> ( Expr )
     *       | CellRef
     *       | CellRange
     *       | Number
     *       | Function
     * @throws \Exception
     * @return array The parsed ptg'd tree
     */
    protected function fact()
    {
        if ($this->currentToken == Token::TOKEN_OPEN) {
            $this->advance(); // eat the "("

            $result = $this->parenthesizedExpression();
            if ($this->currentToken != Token::TOKEN_CLOSE) {
                throw new \Exception("')' token expected.");
            }

            $this->advance(); // eat the ")"

            return $result;
        }

        if (Token::isAnyReference($this->currentToken)) {
            $result = $this->createTree($this->currentToken, '', '');
            $this->advance();

            return $result;
        } elseif (Token::isAnyRange($this->currentToken)) {
            $result = $this->currentToken;
            $this->advance();

            return $result;
        } elseif (is_numeric($this->currentToken)) {
            $result = $this->createTree($this->currentToken, '', '');
            $this->advance();

            return $result;
        } elseif (Token::isFunctionCall($this->currentToken)) {
            $result = $this->func();

            return $result;
        }

        throw new \Exception(
            "Syntax error: " . $this->currentToken .
            ", lookahead: " . $this->lookahead .
            ", current char: " . $this->currentChar
        );
    }

    /**
     * It parses a function call. It assumes the following rule:
     * Func -> ( Expr [,Expr]* )
     * @throws \Exception
     * @return string|array The parsed ptg'd tree
     */
    protected function func()
    {
        $numArgs = 0; // number of arguments received
        $function = strtoupper($this->currentToken);
        $result = ''; // initialize result

        $this->advance();
        $this->advance(); // eat the "("

        while ($this->currentToken != ')') {
            if ($numArgs > 0) {
                if (!Token::isCommaOrSemicolon($this->currentToken)) {
                    throw new \Exception(
                        "Syntax error: comma expected in " .
                        "function $function, arg #{$numArgs}"
                    );
                }

                $this->advance(); // eat the "," or ";"
            } else {
                $result = '';
            }

            $result = $this->createTree('arg', $result, $this->condition());

            $numArgs++;
        }

        $args = Functions::getArgsNumber($function);
        if ($args >= 0 && $args != $numArgs) {
            // If fixed number of args eg. TIME($i,$j,$k). Check that the number of args is valid.
            throw new \Exception("Incorrect number of arguments in function $function() ");
        }

        $result = $this->createTree($function, $result, $numArgs);
        $this->advance(); // eat the ")"

        return $result;
    }

    /**
     * Creates a tree. In fact an array which may have one or two arrays (sub-trees)
     * as elements.
     *
     * @param mixed $value The value of this node.
     * @param mixed $left  The left array (sub-tree) or a final node.
     * @param mixed $right The right array (sub-tree) or a final node.
     * @return array A tree
     */
    protected function createTree($value, $left, $right)
    {
        return array(
            'value' => $value,
            'left' => $left,
            'right' => $right
        );
    }

    /**
     * Builds a string containing the tree in reverse polish notation (What you
     * would use in a HP calculator stack).
     * The following tree:
     *
     *    +
     *   / \
     *  2   3
     *
     * produces: "23+"
     *
     * The following tree:
     *
     *    +
     *   / \
     *  3   *
     *     / \
     *    6   A1
     *
     * produces: "36A1*+"
     *
     * In fact all operands, functions, references, etc... are written as ptg's
     *
     * @param array $tree The optional tree to convert.
     * @return string The tree in reverse polish notation
     */
    protected function toReversePolish($tree)
    {
        if (!is_array($tree)) {
            return $this->convert($tree);
        }

        // if it's a function convert it here (so we can set it's arguments)
        if ($this->isFunction($tree['value'])) {
            return $this->getFunctionPolish($tree);
        }

        $polish = $this->getTreePartPolish($tree['left']);
        $polish .= $this->getTreePartPolish($tree['right']);
        $polish .= $this->convert($tree['value']);

        return $polish;
    }

    /**
     * @param $tree
     *
     * @return string
     */
    protected function getFunctionPolish($tree)
    {
        $polish = '';

        // left subtree for a function is always an array.
        if ($tree['left'] != '') {
            $polish .= $this->toReversePolish($tree['left']);
        }

        $polish .= $this->convertFunction($tree['value'], $tree['right']);

        return $polish;
    }

    /**
     * @param $value
     *
     * @return bool
     */
    protected function isFunction($value)
    {
        return Token::isFunctionCall($value)
            && !Token::isReference($value)
            && !Token::isRangeWithDots($value)
            && !is_numeric($value)
            && !Ptg::exists($value);
    }

    /**
     * @param $part
     *
     * @return string
     * @throws \Exception
     */
    protected function getTreePartPolish($part)
    {
        $polish = '';

        if (is_array($part)) {
            $polish .= $this->toReversePolish($part);
        } elseif ($part != '') {
            // It's a final node
            $polish .= $this->convert($part);
        }

        return $polish;
    }

    /**
     * @return array
     */
    public function getReferences()
    {
        return $this->references;
    }

    /**
     * @param $formula
     *
     * @return string
     */
    public function getReversePolish($formula)
    {
        $this->parse($formula);

        return $this->toReversePolish($this->parseTree);
    }
}
