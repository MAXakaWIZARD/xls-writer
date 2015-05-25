<?php

namespace Xls;

class Ptg
{
    /**
     * @var null|array
     */
    protected static $ptgs = null;

    /**
     * @param $ptg
     *
     * @return bool
     */
    public static function exists($ptg)
    {
        self::cache();

        return isset(self::$ptgs[$ptg]);
    }

    /**
     * @param $ptg
     *
     * @return bool
     */
    public static function get($ptg)
    {
        self::cache();

        return isset(self::$ptgs[$ptg]) ? self::$ptgs[$ptg] : null;
    }

    /**
     *
     */
    protected static function cache()
    {
        if (is_null(self::$ptgs)) {
            self::$ptgs = self::getAll();
        }
    }

    /**
     * @return array
     */
    public static function getAll()
    {
        return array(
            'ptgExp' => 0x01,
            'ptgTbl' => 0x02,
            'ptgAdd' => 0x03,
            'ptgSub' => 0x04,
            'ptgMul' => 0x05,
            'ptgDiv' => 0x06,
            'ptgPower' => 0x07,
            'ptgConcat' => 0x08,
            'ptgLT' => 0x09,
            'ptgLE' => 0x0A,
            'ptgEQ' => 0x0B,
            'ptgGE' => 0x0C,
            'ptgGT' => 0x0D,
            'ptgNE' => 0x0E,
            'ptgIsect' => 0x0F,
            'ptgUnion' => 0x10,
            'ptgRange' => 0x11,
            'ptgUplus' => 0x12,
            'ptgUminus' => 0x13,
            'ptgPercent' => 0x14,
            'ptgParen' => 0x15,
            'ptgMissArg' => 0x16,
            'ptgStr' => 0x17,
            'ptgAttr' => 0x19,
            'ptgSheet' => 0x1A,
            'ptgEndSheet' => 0x1B,
            'ptgErr' => 0x1C,
            'ptgBool' => 0x1D,
            'ptgInt' => 0x1E,
            'ptgNum' => 0x1F,
            'ptgArray' => 0x20,
            'ptgFunc' => 0x21,
            'ptgFuncVar' => 0x22,
            'ptgName' => 0x23,
            'ptgRef' => 0x24,
            'ptgArea' => 0x25,
            'ptgMemArea' => 0x26,
            'ptgMemErr' => 0x27,
            'ptgMemNoMem' => 0x28,
            'ptgMemFunc' => 0x29,
            'ptgRefErr' => 0x2A,
            'ptgAreaErr' => 0x2B,
            'ptgRefN' => 0x2C,
            'ptgAreaN' => 0x2D,
            'ptgMemAreaN' => 0x2E,
            'ptgNameX' => 0x39,
            'ptgRef3d' => 0x3A,
            'ptgArea3d' => 0x3B,
            'ptgRefErr3d' => 0x3C,
            'ptgArrayV' => 0x40,
            'ptgFuncV' => 0x41,
            'ptgFuncVarV' => 0x42,
            'ptgNameV' => 0x43,
            'ptgRefV' => 0x44,
            'ptgAreaV' => 0x45,
            'ptgMemAreaV' => 0x46,
            'ptgMemErrV' => 0x47,
            'ptgMemNoMemV' => 0x48,
            'ptgMemFuncV' => 0x49,
            'ptgRefErrV' => 0x4A,
            'ptgAreaErrV' => 0x4B,
            'ptgRefNV' => 0x4C,
            'ptgAreaNV' => 0x4D,
            'ptgMemAreaNV' => 0x4E,
            'ptgFuncCEV' => 0x58,
            'ptgNameXV' => 0x59,
            'ptgRef3dV' => 0x5A,
            'ptgArea3dV' => 0x5B,
            'ptgRefErr3dV' => 0x5C,
            'ptgArrayA' => 0x60,
            'ptgFuncA' => 0x61,
            'ptgFuncVarA' => 0x62,
            'ptgNameA' => 0x63,
            'ptgRefA' => 0x64,
            'ptgAreaA' => 0x65,
            'ptgMemAreaA' => 0x66,
            'ptgMemErrA' => 0x67,
            'ptgMemNoMemA' => 0x68,
            'ptgMemFuncA' => 0x69,
            'ptgRefErrA' => 0x6A,
            'ptgAreaErrA' => 0x6B,
            'ptgRefNA' => 0x6C,
            'ptgAreaNA' => 0x6D,
            'ptgMemAreaNA' => 0x6E,
            'ptgMemNoMemN' => 0x6F,
            'ptgFuncCEA' => 0x78,
            'ptgNameXA' => 0x79,
            'ptgRef3dA' => 0x7A,
            'ptgArea3dA' => 0x7B,
            'ptgRefErr3dA' => 0x7C,
            'ptgAreaErr3d' => 0x7D
        );
    }
}
