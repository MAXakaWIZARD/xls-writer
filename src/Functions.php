<?php

namespace Xls;

class Functions
{
    /**
     * @var null|array
     */
    protected static $functions = null;

    /**
     * @param $function
     *
     * @return mixed
     * @throws \Exception
     */
    public static function getFunction($function)
    {
        if (is_null(self::$functions)) {
            self::$functions = self::getBuiltIn();
        }

        if (isset(self::$functions[$function])) {
            return self::$functions[$function];
        }

        throw new \Exception("Function $function() doesn't exist");
    }

    /**
     * @param $function
     *
     * @return mixed
     * @throws \Exception
     */
    public static function getArgsNumber($function)
    {
        $function = self::getFunction($function);

        return $function[1];
    }

    /**
     * @param $function
     *
     * @return mixed
     */
    public static function getPtg($function)
    {
        $function = self::getFunction($function);

        return $function[0];
    }

    /**
     * The array elements are as follow:
     *  ptg:   The Excel function ptg code.
     *  args:  The number of arguments that the function takes:
     *           >=0 is a fixed number of arguments.
     *           -1  is a variable  number of arguments.
     *  class: The reference, value or array class of the function args.
     *  vol:   The function is volatile.
     * @return array
     */
    public static function getBuiltIn()
    {
        return array_merge(
            self::getBasicFunctions(),
            self::getTrigonometricFunctions(),
            self::getDatetimeFunctions(),
            self::getMathFunctions(),
            self::getTextFunctions(),
            self::getLogicalFunctions(),
            self::getInformationFunctions(),
            self::getLookupFunctions(),
            self::getStatisticalFunctions(),
            self::getDistributionFunctions(),
            self::getDatabaseFunctions(),
            self::getFinancialFunctions()
        );
    }

    /**
     * @return array
     */
    protected static function getBasicFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'CALL' => array(150, -1, 1, 0),
            'REGISTER.ID' => array(267, -1, 1, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getTrigonometricFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'SIN' => array(15, 1, 1, 0),
            'COS' => array(16, 1, 1, 0),
            'TAN' => array(17, 1, 1, 0),
            'ATAN' => array(18, 1, 1, 0),
            'ATAN2' => array(97, 2, 1, 0),
            'ASIN' => array(98, 1, 1, 0),
            'ACOS' => array(99, 1, 1, 0),
            'RADIANS' => array(342, 1, 1, 0),
            'DEGREES' => array(343, 1, 1, 0),
            'SINH' => array(229, 1, 1, 0),
            'COSH' => array(230, 1, 1, 0),
            'TANH' => array(231, 1, 1, 0),
            'ASINH' => array(232, 1, 1, 0),
            'ACOSH' => array(233, 1, 1, 0),
            'ATANH' => array(234, 1, 1, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getDatetimeFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'DATE' => array(65, 3, 1, 0),
            'TIME' => array(66, 3, 1, 0),
            'DAY' => array(67, 1, 1, 0),
            'MONTH' => array(68, 1, 1, 0),
            'YEAR' => array(69, 1, 1, 0),
            'WEEKDAY' => array(70, -1, 1, 0),
            'HOUR' => array(71, 1, 1, 0),
            'MINUTE' => array(72, 1, 1, 0),
            'SECOND' => array(73, 1, 1, 0),
            'NOW' => array(74, 0, 1, 1),
            'DATEVALUE' => array(140, 1, 1, 0),
            'TIMEVALUE' => array(141, 1, 1, 0),
            'DAYS360' => array(220, -1, 1, 0),
            'TODAY' => array(221, 0, 1, 1),
        );
    }

    /**
     * @return array
     */
    protected static function getMathFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'SUM' => array(4, -1, 0, 0),
            'SUMIF' => array(345, -1, 0, 0),
            'SUMPRODUCT' => array(228, -1, 2, 0),
            'SUMXMY2' => array(303, 2, 2, 0),
            'SUMX2MY2' => array(304, 2, 2, 0),
            'SUMX2PY2' => array(305, 2, 2, 0),
            'SUMSQ' => array(321, -1, 0, 0),
            'PI' => array(19, 0, 1, 0),
            'SQRT' => array(20, 1, 1, 0),
            'EXP' => array(21, 1, 1, 0),
            'LN' => array(22, 1, 1, 0),
            'LOG' => array(109, -1, 1, 0),
            'LOG10' => array(23, 1, 1, 0),
            'ABS' => array(24, 1, 1, 0),
            'INT' => array(25, 1, 1, 0),
            'SIGN' => array(26, 1, 1, 0),
            'ROUND' => array(27, 2, 1, 0),
            'ROUNDUP' => array(212, 2, 1, 0),
            'ROUNDDOWN' => array(213, 2, 1, 0),
            'TRUNC' => array(197, -1, 1, 0),
            'CEILING' => array(288, 2, 1, 0),
            'FLOOR' => array(285, 2, 1, 0),
            'SUBTOTAL' => array(344, -1, 0, 0),
            'POWER' => array(337, 2, 1, 0),
            'MOD' => array(39, 2, 1, 0),
            'PRODUCT' => array(183, -1, 0, 0),
            'ODD' => array(298, 1, 1, 0),
            'EVEN' => array(279, 1, 1, 0),
            'FACT' => array(184, 1, 1, 0),
            'ROMAN' => array(354, -1, 1, 0),
            'COMBIN' => array(276, 2, 1, 0),
            'RAND' => array(63, 0, 1, 1),
            'MDETERM' => array(163, 1, 2, 0),
            'MINVERSE' => array(164, 1, 2, 0),
            'MMULT' => array(165, 2, 2, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getTextFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'CHAR' => array(111, 1, 1, 0),
            'LOWER' => array(112, 1, 1, 0),
            'UPPER' => array(113, 1, 1, 0),
            'PROPER' => array(114, 1, 1, 0),
            'LEFT' => array(115, -1, 1, 0),
            'RIGHT' => array(116, -1, 1, 0),
            'EXACT' => array(117, 2, 1, 0),
            'TRIM' => array(118, 1, 1, 0),
            'REPLACE' => array(119, 4, 1, 0),
            'SUBSTITUTE' => array(120, -1, 1, 0),
            'CLEAN' => array(162, 1, 1, 0),
            'REPT' => array(30, 2, 1, 0),
            'MID' => array(31, 3, 1, 0),
            'CONCATENATE' => array(336, -1, 1, 0),
            'FIND' => array(124, -1, 1, 0),
            'T' => array(130, 1, 0, 0),
            'LEN' => array(32, 1, 1, 0),
            'SEARCH' => array(82, -1, 1, 0),
            'CODE' => array(121, 1, 1, 0),
            'DOLLAR' => array(13, -1, 1, 0),
            'FIXED' => array(14, -1, 1, 0),
            'TEXT' => array(48, 2, 1, 0),
            'VALUE' => array(33, 1, 1, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getLogicalFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'TRUE' => array(34, 0, 1, 0),
            'FALSE' => array(35, 0, 1, 0),
            'AND' => array(36, -1, 0, 0),
            'OR' => array(37, -1, 0, 0),
            'NOT' => array(38, 1, 1, 0),
            'IF' => array(1, -1, 1, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getInformationFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'ISNA' => array(2, 1, 1, 0),
            'ISERROR' => array(3, 1, 1, 0),
            'ISREF' => array(105, 1, 0, 0),
            'ISERR' => array(126, 1, 1, 0),
            'ISTEXT' => array(127, 1, 1, 0),
            'ISNUMBER' => array(128, 1, 1, 0),
            'ISBLANK' => array(129, 1, 1, 0),
            'ISNONTEXT' => array(190, 1, 1, 0),
            'ISLOGICAL' => array(198, 1, 1, 0),
            'CELL' => array(125, -1, 0, 1),
            'INFO' => array(244, 1, 1, 1),
            'NA' => array(10, 0, 0, 0),
            'N' => array(131, 1, 0, 0),
            'TYPE' => array(86, 1, 1, 0),
            'ERROR.TYPE' => array(261, 1, 1, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getLookupFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'LOOKUP' => array(28, -1, 0, 0),
            'HLOOKUP' => array(101, -1, 0, 0),
            'VLOOKUP' => array(102, -1, 0, 0),
            'CHOOSE' => array(100, -1, 1, 0),
            'MATCH' => array(64, -1, 0, 0),
            'AREAS' => array(75, 1, 0, 1),
            'ROW' => array(8, -1, 0, 0),
            'COLUMN' => array(9, -1, 0, 0),
            'ROWS' => array(76, 1, 0, 1),
            'COLUMNS' => array(77, 1, 0, 1),
            'ADDRESS' => array(219, -1, 1, 0),
            'INDEX' => array(29, -1, 0, 1),
            'OFFSET' => array(78, -1, 0, 1),
            'INDIRECT' => array(148, -1, 1, 1),
            'TRANSPOSE' => array(83, 1, 1, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getStatisticalFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'AVERAGE' => array(5, -1, 0, 0),
            'MIN' => array(6, -1, 0, 0),
            'MAX' => array(7, -1, 0, 0),
            'COUNTA' => array(169, -1, 0, 0),
            'COUNT' => array(0, -1, 0, 0),
            'COUNTIF' => array(346, 2, 0, 0),
            'COUNTBLANK' => array(347, 1, 0, 0),
            'STDEVP' => array(193, -1, 0, 0),
            'FREQUENCY' => array(252, 2, 0, 0),
            'MEDIAN' => array(227, -1, 0, 0),
            'MODE' => array(330, -1, 2, 0),
            'TRIMMEAN' => array(331, 2, 0, 0),
            'QUARTILE' => array(327, 2, 0, 0),
            'PERCENTILE' => array(328, 2, 0, 0),
            'PERCENTRANK' => array(329, -1, 0, 0),
            'RANK' => array(216, -1, 0, 0),
            'GEOMEAN' => array(319, -1, 0, 0),
            'HARMEAN' => array(320, -1, 0, 0),
            'KURT' => array(322, -1, 0, 0),
            'STDEV' => array(12, -1, 0, 0),
            'AVEDEV' => array(269, -1, 0, 0),
            'DEVSQ' => array(318, -1, 0, 0),
            'SKEW' => array(323, -1, 0, 0),
            'VAR' => array(46, -1, 0, 0),
            'VARP' => array(194, -1, 0, 0),
            'COVAR' => array(308, 2, 2, 0),
            'LARGE' => array(325, 2, 0, 0),
            'SMALL' => array(326, 2, 0, 0),
            'STEYX' => array(314, 2, 2, 0),
            'CORREL' => array(307, 2, 2, 0),
            'FORECAST' => array(309, 3, 2, 0),
            'SLOPE' => array(315, 2, 2, 0),
            'INTERCEPT' => array(311, 2, 2, 0),
            'TTEST' => array(316, 4, 2, 0),
            'LINEST' => array(49, -1, 0, 0),
            'TREND' => array(50, -1, 0, 0),
            'GROWTH' => array(52, -1, 0, 0),
            'LOGEST' => array(51, -1, 0, 0),
            'PERMUT' => array(299, 2, 1, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getDistributionFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'CONFIDENCE' => array(277, 3, 1, 0),
            'CRITBINOM' => array(278, 3, 1, 0),
            'EXPONDIST' => array(280, 3, 1, 0),
            'FDIST' => array(281, 3, 1, 0),
            'FINV' => array(282, 3, 1, 0),
            'FISHER' => array(283, 1, 1, 0),
            'FISHERINV' => array(284, 1, 1, 0),
            'GAMMADIST' => array(286, 4, 1, 0),
            'GAMMAINV' => array(287, 3, 1, 0),
            'HYPGEOMDIST' => array(289, 4, 1, 0),
            'LOGNORMDIST' => array(290, 3, 1, 0),
            'LOGINV' => array(291, 3, 1, 0),
            'NEGBINOMDIST' => array(292, 3, 1, 0),
            'NORMDIST' => array(293, 4, 1, 0),
            'NORMSDIST' => array(294, 1, 1, 0),
            'NORMINV' => array(295, 3, 1, 0),
            'NORMSINV' => array(296, 1, 1, 0),
            'BETADIST' => array(270, -1, 1, 0),
            'GAMMALN' => array(271, 1, 1, 0),
            'BETAINV' => array(272, -1, 1, 0),
            'BINOMDIST' => array(273, 4, 1, 0),
            'CHIDIST' => array(274, 2, 1, 0),
            'CHIINV' => array(275, 2, 1, 0),
            'STANDARDIZE' => array(297, 3, 1, 0),
            'POISSON' => array(300, 3, 1, 0),
            'TDIST' => array(301, 3, 1, 0),
            'WEIBULL' => array(302, 4, 1, 0),
            'CHITEST' => array(306, 2, 2, 0),
            'FTEST' => array(310, 2, 2, 0),
            'PEARSON' => array(312, 2, 2, 0),
            'RSQ' => array(313, 2, 2, 0),
            'PROB' => array(317, -1, 2, 0),
            'ZTEST' => array(324, -1, 0, 0),
            'TINV' => array(332, 2, 1, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getDatabaseFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'DCOUNT' => array(40, 3, 0, 0),
            'DSUM' => array(41, 3, 0, 0),
            'DAVERAGE' => array(42, 3, 0, 0),
            'DMIN' => array(43, 3, 0, 0),
            'DMAX' => array(44, 3, 0, 0),
            'DSTDEV' => array(45, 3, 0, 0),
            'DVAR' => array(47, 3, 0, 0),
            'DPRODUCT' => array(189, 3, 0, 0),
            'DSTDEVP' => array(195, 3, 0, 0),
            'DVARP' => array(196, 3, 0, 0),
            'DCOUNTA' => array(199, 3, 0, 0),
            'DGET' => array(235, 3, 0, 0),
        );
    }

    /**
     * @return array
     */
    protected static function getFinancialFunctions()
    {
        return array(
            // function ptg  args  class  vol
            'NPV' => array(11, -1, 1, 0),
            'PV' => array(56, -1, 1, 0),
            'FV' => array(57, -1, 1, 0),
            'MIRR' => array(61, 3, 0, 0),
            'IRR' => array(62, -1, 0, 0),
            'PMT' => array(59, -1, 1, 0),
            'IPMT' => array(167, -1, 1, 0),
            'PPMT' => array(168, -1, 1, 0),
            'DDB' => array(144, -1, 1, 0),
            'VDB' => array(222, -1, 1, 0),
            'DB' => array(247, -1, 1, 0),
            'SLN' => array(142, 3, 1, 0),
            'SYD' => array(143, 4, 1, 0),
            'RATE' => array(60, -1, 1, 0),
            'NPER' => array(58, -1, 1, 0),
        );
    }
}
