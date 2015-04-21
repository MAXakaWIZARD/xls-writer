<?php
/* vim: set expandtab tabstop=4 shiftwidth=4: */
// +----------------------------------------------------------------------+
// | PHP Version 4                                                        |
// +----------------------------------------------------------------------+
// | Copyright (c) 1997-2002 The PHP Group                                |
// +----------------------------------------------------------------------+
// | This source file is subject to version 2.02 of the PHP license,      |
// | that is bundled with this package in the file LICENSE, and is        |
// | available at through the world-wide-web at                           |
// | http://www.php.net/license/2_02.txt.                                 |
// | If you did not receive a copy of the PHP license and are unable to   |
// | obtain it through the world-wide-web, please send a note to          |
// | license@php.net so we can mail you a copy immediately.               |
// +----------------------------------------------------------------------+
// | Author: Xavier Noguer <xnoguer@php.net>                              |
// | Based on OLE::Storage_Lite by Kawai, Takanori                        |
// +----------------------------------------------------------------------+
//
// $Id$

namespace Xls\OLE;

use Xls\OLE;

/**
 * Class for creating PPS's for OLE containers
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 * @category Structures
 * @package  OLE
 */
class PPS
{
    /**
     * The PPS index
     * @var integer
     */
    public $No;

    /**
     * The PPS name (in Unicode)
     * @var string
     */
    public $Name;

    /**
     * The PPS type. Dir, Root or File
     * @var integer
     */
    public $Type;

    /**
     * The index of the previous PPS
     * @var integer
     */
    public $PrevPps;

    /**
     * The index of the next PPS
     * @var integer
     */
    public $NextPps;

    /**
     * The index of it's first child if this is a Dir or Root PPS
     * @var integer
     */
    public $DirPps;

    /**
     * A timestamp
     * @var integer
     */
    public $Time1st;

    /**
     * A timestamp
     * @var integer
     */
    public $Time2nd;

    /**
     * Starting block (small or big) for this PPS's data  inside the container
     * @var integer
     */
    public $StartBlock;

    /**
     * The size of the PPS's data (in bytes)
     * @var integer
     */
    public $Size;

    /**
     * The PPS's data (only used if it's not using a temporary file)
     * @var string
     */
    public $data;

    /**
     * Array of child PPS's (only used by Root and Dir PPS's)
     * @var array
     */
    public $children = array();

    /**
     * Pointer to OLE container
     * @var \Xls\OLE
     */
    public $ole;

    /**
     * The constructor
     *
     * @access public
     * @param integer $No   The PPS index
     * @param string $name The PPS name
     * @param integer $type The PPS type. Dir, Root or File
     * @param integer $prev The index of the previous PPS
     * @param integer $next The index of the next PPS
     * @param integer $dir  The index of it's first child if this is a Dir or Root PPS
     * @param integer $time_1st A timestamp
     * @param integer $time_2nd A timestamp
     * @param string $data  The (usually binary) source data of the PPS
     * @param array $children Array containing children PPS for this PPS
     */
    public function __construct(
        $No = null,
        $name = null,
        $type = null,
        $prev = null,
        $next = null,
        $dir = null,
        $time_1st = null,
        $time_2nd = null,
        $data = '',
        $children = array()
    ) {
        $this->No = $No;
        $this->Name = $name;
        $this->Type = $type;
        $this->PrevPps = $prev;
        $this->NextPps = $next;
        $this->DirPps = $dir;
        $this->Time1st = $time_1st;
        $this->Time2nd = $time_2nd;
        $this->data = $data;
        $this->children = $children;

        if ($data != '') {
            $this->Size = strlen($data);
        } else {
            $this->Size = 0;
        }
    }

    /**
     * Returns the amount of data saved for this PPS
     *
     * @access private
     * @return integer The amount of data (in bytes)
     */
    public function dataLen()
    {
        if (!isset($this->data)) {
            return 0;
        }
        if (isset($this->ppsFile)) {
            fseek($this->ppsFile, 0);
            $stats = fstat($this->ppsFile);
            return $stats[7];
        } else {
            return strlen($this->data);
        }
    }

    /**
     * Returns a string with the PPS's WK (What is a WK?)
     *
     * @return string The binary string
     */
    public function getPpsWk()
    {
        $ret = $this->Name;
        for ($i = 0; $i < (64 - strlen($this->Name)); $i++) {
            $ret .= "\x00";
        }
        $ret .= pack("v", strlen($this->Name) + 2) // 66
            . pack("c", $this->Type) // 67
            . pack("c", 0x00) //UK                // 68
            . pack("V", $this->PrevPps) //Prev    // 72
            . pack("V", $this->NextPps) //Next    // 76
            . pack("V", $this->DirPps) //Dir     // 80
            . "\x00\x09\x02\x00" // 84
            . "\x00\x00\x00\x00" // 88
            . "\xc0\x00\x00\x00" // 92
            . "\x00\x00\x00\x46" // 96 // Seems to be ok only for Root
            . "\x00\x00\x00\x00" // 100
            . OLE::localDate2OLE($this->Time1st) // 108
            . OLE::localDate2OLE($this->Time2nd) // 116
            . pack(
                "V",
                isset($this->StartBlock) ? $this->StartBlock : 0
            ) // 120
            . pack("V", $this->Size) // 124
            . pack("V", 0); // 128

        return $ret;
    }

    /**
     * Updates index and pointers to previous, next and children PPS's for this
     * PPS. I don't think it'll work with Dir PPS's.
     *
     * @param array &$raList Reference to the array of PPS's for the whole OLE container
     * @param $toSave
     * @param $depth
     * @return integer          The index for this PPS
     */
    public static function savePpsSetPnt(&$raList, $toSave, $depth = 0)
    {
        if (!is_array($toSave) || (count($toSave) == 0)) {
            return 0xFFFFFFFF;
        } elseif (count($toSave) == 1) {
            $cnt = count($raList);
            // If the first entry, it's the root... Don't clone it!
            $raList[$cnt] = ($depth == 0) ? $toSave[0] : clone $toSave[0];
            $raList[$cnt]->No = $cnt;
            $raList[$cnt]->PrevPps = 0xFFFFFFFF;
            $raList[$cnt]->NextPps = 0xFFFFFFFF;
            $raList[$cnt]->DirPps = self::savePpsSetPnt($raList, @$raList[$cnt]->children, $depth++);

            return $cnt;
        } else {
            $iPos = (int) floor(count($toSave) / 2);
            $aPrev = array_slice($toSave, 0, $iPos);
            $aNext = array_slice($toSave, $iPos + 1);

            $cnt = count($raList);
            // If the first entry, it's the root... Don't clone it!
            $raList[$cnt] = ($depth == 0) ? $toSave[$iPos] : clone $toSave[$iPos];
            $raList[$cnt]->No = $cnt;
            $raList[$cnt]->PrevPps = self::savePpsSetPnt($raList, $aPrev, $depth++);
            $raList[$cnt]->NextPps = self::savePpsSetPnt($raList, $aNext, $depth++);
            $raList[$cnt]->DirPps = self::savePpsSetPnt($raList, @$raList[$cnt]->children, $depth++);

            return $cnt;
        }
    }
}
