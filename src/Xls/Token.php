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

    /**
     * Reference A1 or $A$1
     * @param $token
     *
     * @return boolean
     */
    public static function isReference($token)
    {
        return preg_match('/^\$?[A-Ia-i]?[A-Za-z]\$?[0-9]+$/', $token) === 1;
    }

    /**
     * External reference Sheet1!A1 or Sheet1:Sheet2!A1 or 'Sheet1'!A1 or 'Sheet1:Sheet2'!A1
     * @param $token
     *
     * @return boolean
     */
    public static function isExternalReference($token)
    {
        return preg_match("/^\w+(\:\w+)?\![A-Ia-i]?[A-Za-z][0-9]+$/u", $token) === 1
            || preg_match("/^'[\w -]+(\:[\w -]+)?'\![A-Ia-i]?[A-Za-z][0-9]+$/u", $token) === 1;
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
     * Range A1:A2
     * @param $token
     *
     * @return boolean
     */
    public static function isRangeWithColon($token)
    {
        return preg_match("/^(\$)?[A-Ia-i]?[A-Za-z](\$)?[0-9]+:(\$)?[A-Ia-i]?[A-Za-z](\$)?[0-9]+$/", $token) === 1;
    }

    /**
     * Range A1..A2
     * @param $token
     *
     * @return boolean
     */
    public static function isRangeWithDots($token)
    {
        return preg_match("/^(\$)?[A-Ia-i]?[A-Za-z](\$)?[0-9]+\.\.(\$)?[A-Ia-i]?[A-Za-z](\$)?[0-9]+$/", $token) === 1;
    }

    /**
     * External range Sheet1!A1 or Sheet1:Sheet2!A1:B2 or 'Sheet1'!A1 or 'Sheet1:Sheet2'!A1:B2
     * @param $token
     *
     * @return boolean
     */
    public static function isExternalRange($token)
    {
        return preg_match(
            "/^\w+(\:\w+)?\!([A-Ia-i]?[A-Za-z])?[0-9]+:([A-Ia-i]?[A-Za-z])?[0-9]+$/u",
            $token
        ) === 1
        || preg_match(
            "/^'[\w -]+(\:[\w -]+)?'\!([A-Ia-i]?[A-Za-z])?[0-9]+:([A-Ia-i]?[A-Za-z])?[0-9]+$/u",
            $token
        ) === 1;
    }

    /**
     * String (of maximum 255 characters)
     * @param $token
     *
     * @return boolean
     */
    public static function isString($token)
    {
        return preg_match("/^\"[^\"]{0,255}\"$/", $token) === 1;
    }

    /**
     * @param $token
     *
     * @return boolean
     */
    public static function isFunctionCall($token)
    {
        return preg_match("/^[A-Z0-9\xc0-\xdc\.]+$/", $token) === 1;
    }
}
