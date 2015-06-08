<?php

namespace Xls\OLE;

class PpsRoot extends PPS
{
    /**
     * @var
     */
    protected $smallBlockSize;

    /**
     * @var
     */
    protected $bigBlockSize;

    /**
     * @var
     */
    protected $rootFilePointer;

    /**
     * @param integer $timestamp A timestamp
     * @param PpsFile[] $children
     */
    public function __construct(
        $timestamp = null,
        $children = array()
    ) {
        parent::__construct(
            null,
            OLE::asc2Ucs('Root Entry'),
            self::PPS_TYPE_ROOT,
            null,
            null,
            null,
            $timestamp,
            null,
            $children
        );
    }

    /**
     * Method for saving the whole OLE container (including files).
     *
     * @param string $filename The name of the file where to save the OLE container
     * @throws \Exception
     * @return boolean
     */
    public function save($filename)
    {
        $this->openFile($filename);

        $this->setBlockSizes();

        // Make an array of PPS's (for Save)
        $list = array();
        self::setPointers($list, array($this));

        list($iSBDcnt, $iBBcnt, $iPPScnt) = $this->calcSize($list);

        $this->saveHeader($iSBDcnt, $iBBcnt, $iPPScnt);
        $this->saveSmallData($list);
        $this->saveBigData($iSBDcnt, $list);
        $this->savePps($list);
        $this->saveBigBlockChain($iSBDcnt, $iBBcnt, $iPPScnt);

        fclose($this->rootFilePointer);
    }

    /**
     *
     */
    protected function setBlockSizes()
    {
        $this->bigBlockSize = pow(2, 9);
        $this->smallBlockSize = pow(2, 6);
    }

    /**
     * @param $filename
     *
     * @throws \Exception
     */
    protected function openFile($filename)
    {
        $this->rootFilePointer = @fopen($filename, "wb");
        if ($this->rootFilePointer === false) {
            throw new \Exception("Can't open $filename. It may be in use or protected.");
        }
    }

    /**
     * @param $size
     * @param $blockSize
     *
     * @return float
     */
    protected function getBlocksCount($size, $blockSize)
    {
        return floor($size / $blockSize) + (($size % $blockSize) ? 1 : 0);
    }

    /**
     * Calculate some numbers
     *
     * @param PPS[] $list Reference to an array of PPS's
     *
     * @return array The array of numbers
     */
    protected function calcSize($list)
    {
        $iSBcnt = 0;
        $iBBcnt = 0;
        foreach ($list as $item) {
            if (!$item->isFile()) {
                continue;
            }

            $size = $item->getSize();
            if ($size < self::DATA_SIZE_SMALL) {
                $iSBcnt += $this->getBlocksCount($size, $this->smallBlockSize);
            } else {
                $iBBcnt += $this->getBlocksCount($size, $this->bigBlockSize);
            }
        }

        $iSmallLen = $iSBcnt * $this->smallBlockSize;
        $iSlCnt = $this->getPointersPerBlock($this->bigBlockSize);
        $iSBDcnt = $this->getBlocksCount($iSBcnt, $iSlCnt);
        $iBBcnt += $this->getBlocksCount($iSmallLen, $this->bigBlockSize);
        $iBdCnt = $this->getPointersPerBlock($this->bigBlockSize, self::PPS_SIZE);
        $iPPScnt = $this->getBlocksCount(count($list), $iBdCnt);

        return array($iSBDcnt, $iBBcnt, $iPPScnt);
    }

    /**
     * Saving big data (PPS's with data bigger than OLE_DATA_SIZE_SMALL)
     *
     * @param integer $iStBlk
     * @param PPS[] &$list Reference to array of PPS's
     */
    public function saveBigData($iStBlk, &$list)
    {
        foreach ($list as $item) {
            $size = $item->getSize();
            if ($size >= self::DATA_SIZE_SMALL
                || ($item->isRoot() && $item->hasData())
            ) {
                $this->write($item->getData());

                if ($size % $this->bigBlockSize) {
                    $zeroByteCount = ($this->bigBlockSize - ($size % $this->bigBlockSize));
                    $this->write(str_repeat("\x00", $zeroByteCount));
                }

                // Set For PPS
                $item->setStartBlock($iStBlk);
                $iStBlk += $this->getBlocksCount($size, $this->bigBlockSize);
            }
        }
    }

    /**
     * get small data (PPS's with data smaller than OLE_DATA_SIZE_SMALL)
     *
     * @param PPS[] &$list Reference to array of PPS's
     * @return string
     */
    protected function saveSmallData(&$list)
    {
        $result = '';
        $iSmBlk = 0;
        foreach ($list as $item) {
            if (!$item->isFile()) {
                continue;
            }

            $size = $item->getSize();
            if ($size <= 0 || $size >= self::DATA_SIZE_SMALL) {
                continue;
            }

            $sbCount = $this->getBlocksCount($size, $this->smallBlockSize);
            // Add to SBD
            for ($j = 0; $j < $sbCount - 1; $j++) {
                $this->writeUlong($j + $iSmBlk + 1);
            }
            $this->writeUlong(-2);

            $result .= $item->getData();

            $exp = $size % $this->smallBlockSize;
            if ($exp) {
                $zeroByteCount = $this->smallBlockSize - $exp;
                $result .= str_repeat("\x00", $zeroByteCount);
            }

            $item->setStartBlock($iSmBlk);
            $iSmBlk += $sbCount;
        }

        $sbCount = $this->getPointersPerBlock($this->bigBlockSize);
        if ($iSmBlk % $sbCount) {
            $repeatCount = $sbCount - ($iSmBlk % $sbCount);
            $this->writeUlong(-1, $repeatCount);
        }

        $this->data = $result;
    }

    /**
     * Saves all the PPS's WKs
     *
     * @param PPS[] $list Reference to an array with all PPS's
     */
    protected function savePps(&$list)
    {
        // Save each PPS WK
        $raListCount = count($list);
        for ($i = 0; $i < $raListCount; $i++) {
            $this->write($list[$i]->getPpsWk());
        }

        // Adjust for Block
        $iCnt = count($list);
        $iBCnt = $this->getPointersPerBlock($this->bigBlockSize, self::PPS_SIZE);
        if ($iCnt % $iBCnt) {
            $zeroByteCount = ($iBCnt - ($iCnt % $iBCnt)) * self::PPS_SIZE;
            $this->write(str_repeat("\x00", $zeroByteCount));
        }
    }

    /**
     * @param     $value
     * @param int $count
     */
    protected function writeUlong($value, $count = 1)
    {
        $packed = pack("V", $value);
        $this->write(str_repeat($packed, $count));
    }

    /**
     * @param $data
     */
    protected function write($data)
    {
        fwrite($this->rootFilePointer, $data);
    }

    /**
     * Saving Big Block Depot
     *
     * @param integer $numSbBlocks - number of Smallblock depot blocks
     * @param integer $numBbBlocks - number of Bigblock depot blocks
     * @param integer $numPpsBlocks - number of PropertySetStorage blocks
     */
    protected function saveBigBlockChain($numSbBlocks, $numBbBlocks, $numPpsBlocks)
    {
        $info = $this->calcBigBlockChain($numSbBlocks, $numBbBlocks, $numPpsBlocks);
        $headerEntriesCount = $info["header_list_entries"];
        $entriesCount = $info["list_entries"];
        $entriesPerBlock = $info["entries_per_block"];

        if ($numSbBlocks > 0) {
            for ($i = 0; $i < $numSbBlocks - 1; $i++) {
                $this->writeUlong($i + 1);
            }
            $this->writeUlong(-2);
        }

        for ($i = 0; $i < $numBbBlocks - 1; $i++) {
            $this->writeUlong($i + $numSbBlocks + 1);
        }
        $this->writeUlong(-2);

        for ($i = 0; $i < $numPpsBlocks - 1; $i++) {
            $this->writeUlong($i + $numSbBlocks + $numBbBlocks + 1);
        }
        $this->writeUlong(-2);

        $this->writeUlong(0xFFFFFFFD, $info["FD_entries"]);
        $this->writeUlong(0xFFFFFFFC, $info["FC_entries"]);

        // Adjust for Block
        $allEntries = $numSbBlocks
            + $numBbBlocks
            + $numPpsBlocks
            + $info["FD_entries"]
            + $info["FC_entries"];
        if ($allEntries % $entriesPerBlock) {
            $rest = $entriesPerBlock - ($allEntries % $entriesPerBlock);
            $this->writeUlong(-1, $rest);
        }

        // Extra BDList
        if ($entriesCount > $headerEntriesCount) {
            $iN = 0;
            $iNb = 0;
            $lastEntryIdx = $entriesPerBlock - 1;

            for ($i = $headerEntriesCount; $i < $entriesCount; $i++, $iN++) {
                if ($iN >= $lastEntryIdx) {
                    $iN = 0;
                    $iNb++;

                    $val = $numSbBlocks
                        + $numBbBlocks
                        + $numPpsBlocks
                        + $info["FD_entries"]
                        + $iNb;
                    $this->writeUlong($val);
                }

                $this->writeUlong($numBbBlocks + $numSbBlocks + $numPpsBlocks + $i);
            }

            $allEntries = $entriesCount - $headerEntriesCount;
            if ($allEntries % $lastEntryIdx) {
                $rest = $lastEntryIdx - ($allEntries % $lastEntryIdx);
                $this->writeUlong(-1, $rest);
            }

            $this->writeUlong(-2);
        }
    }

    /**
     * Save OLE header
     *
     * @param integer $numSbBlocks - number of Smallblock depot blocks
     * @param integer $numBbBlocks - number of Bigblock depot blocks
     * @param integer $numPpsBlocks - number of PropertySetStorage blocks
     */
    public function saveHeader($numSbBlocks, $numBbBlocks, $numPpsBlocks)
    {
        $info = $this->calcBigBlockChain($numSbBlocks, $numBbBlocks, $numPpsBlocks);
        $headerEntriesCount = $info["header_list_entries"];
        $entriesCount = $info["list_entries"];

        $this->write(
            "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
            . pack("V4", 0, 0, 0, 0)
            . pack("v6", 0x3b, 0x03, -2, 9, 6, 0)
            . pack("V2", 0, 0)
            . pack("V", $entriesCount)
            . pack("V", $numSbBlocks + $numBbBlocks) //ROOT START
            . pack("V2", 0, 0x1000)
        );

        //Small Block Depot
        $value = ($numSbBlocks > 0) ? 0 : -2;
        $this->writeUlong($value);
        $this->writeUlong($numSbBlocks);

        // Extra BDList Start, Count
        if ($entriesCount < $headerEntriesCount) {
            $this->writeUlong(-2); // Extra BDList Start
            $this->writeUlong(0); // Extra BDList Count
        } else {
            $this->writeUlong($numSbBlocks + $numBbBlocks + $numPpsBlocks + $info["FD_entries"]);
            $this->writeUlong($info["FC_entries"]);
        }

        // BDList
        for ($i = 0; $i < $headerEntriesCount && $i < $entriesCount; $i++) {
            $this->writeUlong($numBbBlocks + $numSbBlocks + $numPpsBlocks + $i);
        }

        if ($i < $headerEntriesCount) {
            $repeatCount = $headerEntriesCount - $i;
            $this->writeUlong(-1, $repeatCount);
        }
    }

    /**
     * New method to calculate Bigblock chain
     *
     * @param integer $numSb - number of Smallblock depot blocks
     * @param integer $numBb - number of Bigblock depot blocks
     * @param integer $numPps - number of PropertySetStorage blocks
     * @return array
     */
    protected function calcBigBlockChain($numSb, $numBb, $numPps)
    {
        $totalBlocks = $numSb + $numBb + $numPps;
        $info = array(
            "entries_per_block" => $this->getPointersPerBlock($this->bigBlockSize),
            "header_list_entries" => $this->getPointersPerBlock($this->bigBlockSize - 0x4C),
            "entries" => $totalBlocks,
            "ext_list_entries" => 0,
            "FC_entries" => 0,
            "FD_entries" => $this->getNumberOfPointerBlocks($totalBlocks)
        );

        $info["list_entries"] = $this->getNumberOfPointerBlocks(
            $totalBlocks + $info["FD_entries"]
        );

        if ($info["list_entries"] <= $info["header_list_entries"]) {
            return $info;
        }

        return $this->calcBigBlockChainExtra($info);
    }

    /**
     * @param array $info
     *
     * @return array
     */
    protected function calcBigBlockChainExtra($info)
    {
        while (true) {
            $pointerBlocksCount = $this->getNumberOfPointerBlocks(
                $info["entries"] + $info["FD_entries"] + $info["FC_entries"]
            );

            if ($info["list_entries"] >= $pointerBlocksCount) {
                break;
            }

            $info["list_entries"] = $pointerBlocksCount;
            $info["ext_list_entries"] = $info["list_entries"] - $info["header_list_entries"];
            $info["FC_entries"] = $this->getNumberOfPointerBlocks(
                $info["ext_list_entries"]
            );
            $info["FD_entries"] = $this->getNumberOfPointerBlocks(
                $info["entries"] + $info["FD_entries"] + $info["FC_entries"]
            );
        }

        return $info;
    }

    /**
     * Calculates number of pointer blocks
     *
     * @param integer $numPointers - number of pointers
     *
     * @return int
     */
    protected function getNumberOfPointerBlocks($numPointers)
    {
        return $this->getBlocksCount($numPointers, $this->getPointersPerBlock($this->bigBlockSize));
    }

    /**
     * @param int $blockSize
     * @param int $pointerSize
     *
     * @return int
     */
    protected function getPointersPerBlock($blockSize, $pointerSize = self::LONG_INT_SIZE)
    {
        return intval(floor($blockSize / $pointerSize));
    }
}
