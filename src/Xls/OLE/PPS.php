<?php

namespace Xls\OLE;

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
    protected $index;

    /**
     * The PPS name (in Unicode)
     * @var string
     */
    protected $name;

    /**
     * The PPS type. Dir, Root or File
     * @var integer
     */
    protected $type;

    /**
     * The index of the previous PPS
     * @var integer
     */
    protected $prevPps;

    /**
     * The index of the next PPS
     * @var integer
     */
    protected $nextPps;

    /**
     * The index of it's first child if this is a Dir or Root PPS
     * @var integer
     */
    protected $dirPps;

    /**
     * A timestamp
     * @var integer
     */
    protected $time1st;

    /**
     * A timestamp
     * @var integer
     */
    protected $time2nd;

    /**
     * Starting block (small or big) for this PPS's data  inside the container
     * @var integer
     */
    protected $startBlock;

    /**
     * The PPS's data (only used if it's not using a temporary file)
     * @var string
     */
    protected $data;

    /**
     * Array of child PPS's (only used by Root and Dir PPS's)
     * @var array
     */
    protected $children = array();

    /**
     * The temporary dir for storing the OLE file
     * @var string
     */
    protected $tmpDir;

    /**
     * @var string
     */
    protected $tmpFilename;

    /**
     * @var resource
     */
    protected $filePointer;

    /**
     * The constructor
     *
     * @param integer $index The PPS index
     * @param string $name The PPS name
     * @param integer $type The PPS type. Dir, Root or File
     * @param integer $prev The index of the previous PPS
     * @param integer $next The index of the next PPS
     * @param integer $dir  The index of it's first child if this is a Dir or Root PPS
     * @param integer $time1st A timestamp
     * @param integer $time2nd A timestamp
     * @param string $data  The (usually binary) source data of the PPS
     * @param PPS[] $children Array containing children PPS for this PPS
     */
    public function __construct(
        $index = null,
        $name = null,
        $type = null,
        $prev = null,
        $next = null,
        $dir = null,
        $time1st = null,
        $time2nd = null,
        $data = '',
        $children = array()
    ) {
        $this->index = $index;
        $this->name = $name;
        $this->type = $type;
        $this->prevPps = $prev;
        $this->nextPps = $next;
        $this->dirPps = $dir;
        $this->time1st = $time1st;
        $this->time2nd = $time2nd;

        $this->children = $children;

        $this->data = $data;

        $this->tmpDir = sys_get_temp_dir();
    }

    /**
     *
     */
    public function __destruct()
    {
        $this->removeTmpFile();
    }

    /**
     * @return bool
     */
    protected function hasData()
    {
        return isset($this->data);
    }

    /**
     * @return string
     */
    public function getData()
    {
        return $this->data;
    }

    /**
     * Returns the amount of data saved for this PPS
     *
     * @return integer The amount of data (in bytes)
     */
    protected function getSize()
    {
        if (is_resource($this->filePointer)) {
            fseek($this->filePointer, 0);
            $stats = fstat($this->filePointer);
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
        $ret = $this->name;
        for ($i = 0; $i < (64 - strlen($this->name)); $i++) {
            $ret .= "\x00";
        }

        $ret .= pack("v", strlen($this->name) + 2) // 66
            . pack("c", $this->type) // 67
            . pack("c", 0x00) //UK                // 68
            . pack("V", $this->prevPps) //Prev    // 72
            . pack("V", $this->nextPps) //Next    // 76
            . pack("V", $this->dirPps) //Dir     // 80
            . "\x00\x09\x02\x00" // 84
            . "\x00\x00\x00\x00" // 88
            . "\xc0\x00\x00\x00" // 92
            . "\x00\x00\x00\x46" // 96 // Seems to be ok only for Root
            . "\x00\x00\x00\x00" // 100
            . OLE::localDate2OLE($this->time1st) // 108
            . OLE::localDate2OLE($this->time2nd) // 116
            . pack("V", $this->getStartBlock()) // 120
            . pack("V", $this->getSize()) // 124
            . pack("V", 0); // 128

        return $ret;
    }

    /**
     * Updates index and pointers to previous, next and children PPS's for this
     * PPS. I don't think it'll work with Dir PPS's.
     *
     * @param PPS[] &$list Reference to the array of PPS's for the whole OLE container
     * @param array $toSave
     * @param $depth
     *
     * @return integer          The index for this PPS
     */
    public static function savePpsSetPnt(&$list, $toSave, $depth = 0)
    {
        if (!is_array($toSave) || (count($toSave) == 0)) {
            return 0xFFFFFFFF;
        } elseif (count($toSave) == 1) {
            $cnt = count($list);
            // If the first entry, it's the root... Don't clone it!
            $list[$cnt] = ($depth == 0) ? $toSave[0] : clone $toSave[0];
            $list[$cnt]->setIndex($cnt);
            $list[$cnt]->setPrevPps(0xFFFFFFFF);
            $list[$cnt]->setNextPps(0xFFFFFFFF);
            $list[$cnt]->setDirPps(self::savePpsSetPnt($list, $list[$cnt]->getChildren(), $depth++));

            return $cnt;
        } else {
            $iPos = (int)floor(count($toSave) / 2);
            $aPrev = array_slice($toSave, 0, $iPos);
            $aNext = array_slice($toSave, $iPos + 1);

            $cnt = count($list);
            // If the first entry, it's the root... Don't clone it!
            $list[$cnt] = ($depth == 0) ? $toSave[$iPos] : clone $toSave[$iPos];
            $list[$cnt]->setIndex($cnt);
            $list[$cnt]->setPrevPps(self::savePpsSetPnt($list, $aPrev, $depth++));
            $list[$cnt]->setNextPps(self::savePpsSetPnt($list, $aNext, $depth++));
            $list[$cnt]->setDirPps(self::savePpsSetPnt($list, $list[$cnt]->getChildren(), $depth++));

            return $cnt;
        }
    }

    /**
     *
     */
    public function removeTmpFile()
    {
        if (is_resource($this->filePointer)) {
            fclose($this->filePointer);
            $this->filePointer = null;
        }
        @unlink($this->tmpFilename);
    }

    /**
     * @return resource
     */
    public function getFilePointer()
    {
        return $this->filePointer;
    }

    /**
     * @return bool
     */
    public function isFile()
    {
        return $this->type == OLE::PPS_TYPE_FILE;
    }

    /**
     * @return bool
     */
    public function isDir()
    {
        return $this->type == OLE::PPS_TYPE_DIR;
    }

    /**
     * @return bool
     */
    public function isRoot()
    {
        return $this->type == OLE::PPS_TYPE_ROOT;
    }

    /**
     * @param int $startBlock
     */
    public function setStartBlock($startBlock)
    {
        $this->startBlock = $startBlock;
    }

    /**
     * @return int
     */
    public function getStartBlock()
    {
        return isset($this->startBlock) ? $this->startBlock : 0;
    }

    /**
     * @return array
     */
    public function getChildren()
    {
        return $this->children;
    }

    /**
     * @param int $index
     */
    public function setIndex($index)
    {
        $this->index = $index;
    }

    /**
     * @param int $prevPps
     */
    public function setPrevPps($prevPps)
    {
        $this->prevPps = $prevPps;
    }

    /**
     * @param int $nextPps
     */
    public function setNextPps($nextPps)
    {
        $this->nextPps = $nextPps;
    }

    /**
     * @param int $dirPps
     */
    public function setDirPps($dirPps)
    {
        $this->dirPps = $dirPps;
    }

    /**
     * @param $stream
     *
     * @return string
     */
    protected function getStreamContent($stream)
    {
        $content = '';
        fseek($stream, 0);
        while ($buffer = fread($stream, 4096)) {
            $content .= $buffer;
        }

        return $content;
    }
}
