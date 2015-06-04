<?php

namespace Xls;

class Token
{
    const TOKEN_ADD = "+";
    const TOKEN_SUB = "-";
    const TOKEN_MUL = "*";
    const TOKEN_DIV = "/";
    const TOKEN_OPEN = "(";
    const TOKEN_CLOSE = ")";
    const TOKEN_COMA = ",";
    const TOKEN_SEMICOLON = ";";
    const TOKEN_GT = ">";
    const TOKEN_LT = "<";
    const TOKEN_LE = "<=";
    const TOKEN_GE = ">=";
    const TOKEN_EQ = "=";
    const TOKEN_NE = "<>";
    const TOKEN_CONCAT = "&";
    const TOKEN_ARG = "arg";

    protected static $ptgMap = array(
        self::TOKEN_MUL => 'ptgMul',
        self::TOKEN_DIV => 'ptgDiv',
        self::TOKEN_ADD => 'ptgAdd',
        self::TOKEN_SUB => 'ptgSub',
        self::TOKEN_LT => 'ptgLT',
        self::TOKEN_GT => 'ptgGT',
        self::TOKEN_LE => 'ptgLE',
        self::TOKEN_GE => 'ptgGE',
        self::TOKEN_EQ => 'ptgEQ',
        self::TOKEN_NE => 'ptgNE',
        self::TOKEN_CONCAT => 'ptgConcat',
    );

    protected static $deterministicMap = array(
        self::TOKEN_MUL => 1,
        self::TOKEN_DIV => 1,
        self::TOKEN_ADD => 1,
        self::TOKEN_SUB => 1,
        self::TOKEN_LE => 1,
        self::TOKEN_GE => 1,
        self::TOKEN_EQ => 1,
        self::TOKEN_NE => 1,
        self::TOKEN_CONCAT => 1,
        self::TOKEN_COMA => 1,
        self::TOKEN_SEMICOLON => 1,
        self::TOKEN_OPEN => 1,
        self::TOKEN_CLOSE => 1
    );

    protected static $lookaheadMap = array(
        self::TOKEN_GT => array('='),
        self::TOKEN_LT => array('=', '>'),
    );

    protected static $comparisonTokens = array(
        self::TOKEN_LT,
        self::TOKEN_GT,
        self::TOKEN_LE,
        self::TOKEN_GE,
        self::TOKEN_EQ,
        self::TOKEN_NE
    );

    /**
     * Reference A1 or $A$1
     * @param $token
     *
     * @return boolean
     */
    public static function isReference($token)
    {
        return preg_match('/^\$?[A-Za-z]+\$?[0-9]+$/', $token) === 1;
    }

    /**
     * External reference Sheet1!A1 or Sheet1:Sheet2!A1 or 'Sheet1'!A1 or 'Sheet1:Sheet2'!A1
     * @param $token
     *
     * @return boolean
     */
    public static function isExternalReference($token)
    {
        return preg_match("/^\w+(\:\w+)?\![A-za-z]+[0-9]+$/u", $token) === 1
            || preg_match("/^'[\w -]+(\:[\w -]+)?'\![A-za-z]+[0-9]+$/u", $token) === 1;
    }

    /**
     * @param $token
     *
     * @return boolean
     */
    public static function isAnyReference($token)
    {
        return self::isReference($token) || self::isExternalReference($token);
    }

    /**
     * @param $token
     *
     * @return boolean
     */
    public static function isAnyRange($token)
    {
        return self::isRange($token) || self::isExternalRange($token);
    }

    /**
     * Range A1:A2 or A1..A2
     * @param $token
     *
     * @return boolean
     */
    public static function isRange($token)
    {
        return self::isRangeWithColon($token) || self::isRangeWithDots($token);
    }

    /**
     * Range A1:A2 or $A$1:$A$2
     * @param $token
     *
     * @return boolean
     */
    public static function isRangeWithColon($token)
    {
        return preg_match('/^(\$)?[A-Ia-i]?[A-Za-z](\$)?[0-9]+:(\$)?[A-Ia-i]?[A-Za-z](\$)?[0-9]+$/', $token) === 1;
    }

    /**
     * Range A1..A2 or $A$1..$A$2
     * @param $token
     *
     * @return boolean
     */
    public static function isRangeWithDots($token)
    {
        return preg_match('/^(\$)?[A-Ia-i]?[A-Za-z](\$)?[0-9]+\.\.(\$)?[A-Ia-i]?[A-Za-z](\$)?[0-9]+$/', $token) === 1;
    }

    /**
     * External range:
     * Sheet1!A1:B2 or Sheet1:Sheet2!A1:B2
     * 'Sheet1'!A1:B2 or 'Sheet1:Sheet2'!A1:B2
     *
     * @param $token
     *
     * @return boolean
     */
    public static function isExternalRange($token)
    {
        // A1:B2
        $cellsPattern = '([A-Ia-i]?[A-Za-z])?[0-9]+:([A-Ia-i]?[A-Za-z])?[0-9]+';

        // Sheet1!A1:B2 or Sheet1:Sheet2!A1:B2
        $unquotedSheetsPattern = "/^\w+(\:\w+)?\!{$cellsPattern}$/u";

        // 'Sheet1'!A1:B2 or 'Sheet1:Sheet2'!A1:B2
        $quotedSheetsPattern = "/^'[\w -]+(\:[\w -]+)?'\!{$cellsPattern}$/u";

        return preg_match($unquotedSheetsPattern, $token) === 1
            || preg_match($quotedSheetsPattern, $token) === 1;
    }

    /**
     * String
     * @param $token
     *
     * @return boolean
     */
    public static function isString($token)
    {
        return preg_match("/^\"[^\"]*\"$/", $token) === 1;
    }

    /**
     * @param $token
     *
     * @return boolean
     */
    public static function isFunctionCall($token)
    {
        if (self::isArg($token)) {
            return false;
        }

        return preg_match("/^[A-Za-z0-9\xc0-\xdc\.]+$/", $token) === 1;
    }

    /**
     * @param $token
     *
     * @return bool
     */
    public static function isMulOrDiv($token)
    {
        return $token == self::TOKEN_MUL || $token == self::TOKEN_DIV;
    }

    /**
     * @param $token
     *
     * @return bool
     */
    public static function isAddOrSub($token)
    {
        return $token == self::TOKEN_ADD || $token == self::TOKEN_SUB;
    }

    /**
     * @param $token
     *
     * @return bool
     */
    public static function isCommaOrSemicolon($token)
    {
        return $token == self::TOKEN_COMA || $token == self::TOKEN_SEMICOLON;
    }

    /**
     * @param $token
     *
     * @return bool
     */
    public static function isComparison($token)
    {
        return in_array($token, self::$comparisonTokens);
    }

    /**
     * @param $token
     *
     * @return bool
     */
    public static function isLtOrGt($token)
    {
        return $token == self::TOKEN_LT
        || $token == self::TOKEN_GT;
    }

    /**
     * @param $token
     *
     * @return bool
     */
    public static function isConcat($token)
    {
        return $token == self::TOKEN_CONCAT;
    }

    /**
     * @param $token
     *
     * @return string|null
     */
    public static function getPtg($token)
    {
        if (isset(self::$ptgMap[$token])) {
            return self::$ptgMap[$token];
        }

        return null;
    }

    /**
     * @param $token
     *
     * @return bool
     */
    public static function isDeterministic($token)
    {
        return isset(self::$deterministicMap[$token]);
    }

    /**
     * @param $token
     * @param $lookahead
     *
     * @return bool
     */
    public static function isPossibleLookahead($token, $lookahead)
    {
        if (!isset(self::$lookaheadMap[$token])) {
            return true;
        }

        return in_array($lookahead, self::$lookaheadMap[$token], true);
    }

    /**
     * @param $token
     *
     * @return bool
     */
    public static function isArg($token)
    {
        return $token == self::TOKEN_ARG;
    }
}
