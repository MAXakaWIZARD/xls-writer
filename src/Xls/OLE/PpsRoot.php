<?php

namespace Xls\OLE;

/**
 * Class for creating Root PPS's for OLE containers
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 * @category Structures
 * @package  OLE
 */
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
            OLE::PPS_TYPE_ROOT,
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
     * @return boolean true on success
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

        // Make Small Data string (write SBD)
        $this->data = $this->getAndWriteSmallData($list);

        $this->saveBigData($iSBDcnt, $list);
        $this->savePps($list);
        $this->saveBigBlockChain($iSBDcnt, $iBBcnt, $iPPScnt);

        fclose($this->rootFilePointer);

        return true;
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
            if ($size < OLE::DATA_SIZE_SMALL) {
                $iSBcnt += $this->getBlocksCount($size, $this->smallBlockSize);
            } else {
                $iBBcnt += $this->getBlocksCount($size, $this->bigBlockSize);
            }
        }

        $iSmallLen = $iSBcnt * $this->smallBlockSize;
        $iSlCnt = $this->getPointersPerBlock($this->bigBlockSize);
        $iSBDcnt = $this->getBlocksCount($iSBcnt, $iSlCnt);
        $iBBcnt += $this->getBlocksCount($iSmallLen, $this->bigBlockSize);
        $iBdCnt = $this->getPointersPerBlock($this->bigBlockSize, OLE::PPS_SIZE);
        $iPPScnt = $this->getBlocksCount(count($list), $iBdCnt);

        return array($iSBDcnt, $iBBcnt, $iPPScnt);
    }

    /**
     * Saving big data (PPS's with data bigger than OLE_DATA_SIZE_SMALL)
     *
     * @param integer $iStBlk
     * @param PPS[] &$raList Reference to array of PPS's
     */
    public function saveBigData($iStBlk, &$raList)
    {
        foreach ($raList as $item) {
            if ($item->isDir()) {
                continue;
            }

            $size = $item->getSize();
            if ($size >= OLE::DATA_SIZE_SMALL
                || ($item->isRoot() && $item->hasData())
            ) {
                // Write Data
                $filePointer = $item->getFilePointer();
                if (is_resource($filePointer)) {
                    $this->copyFromItemStream($filePointer);
                } else {
                    $this->write($item->getData());
                }

                if ($size % $this->bigBlockSize) {
                    $loopEnd = ($this->bigBlockSize - ($size % $this->bigBlockSize));
                    for ($j = 0; $j < $loopEnd; $j++) {
                        $this->write("\x00");
                    }
                }
                // Set For PPS
                $item->setStartBlock($iStBlk);
                $iStBlk += $this->getBlocksCount($size, $this->bigBlockSize);
            }
        }
    }

    /**
     * @param $sourceStream
     */
    protected function copyFromItemStream($sourceStream)
    {
        fseek($sourceStream, 0);
        while ($sBuff = fread($sourceStream, 4096)) {
            $this->write($sBuff);
        }
    }

    /**
     * get small data (PPS's with data smaller than OLE_DATA_SIZE_SMALL)
     *
     * @param PPS[] &$list Reference to array of PPS's
     * @return string
     */
    protected function getAndWriteSmallData(&$list)
    {
        $result = '';
        $iSmBlk = 0;
        foreach ($list as $item) {
            if (!$item->isFile()) {
                continue;
            }

            $size = $item->getSize();
            if ($size <= 0 || $size >= OLE::DATA_SIZE_SMALL) {
                continue;
            }

            $sbCount = $this->getBlocksCount($size, $this->smallBlockSize);
            // Add to SBD
            for ($j = 0; $j < ($sbCount - 1); $j++) {
                $this->writeUlong($j + $iSmBlk + 1);
            }
            $this->writeUlong(-2);

            // Add to Data String(this will be written for RootEntry)
            $filePointer = $item->getFilePointer();
            if (is_resource($filePointer)) {
                $result .= $this->getStreamContent($filePointer);
            } else {
                $result .= $item->getData();
            }

            $exp = $size % $this->smallBlockSize;
            if ($exp) {
                for ($j = 0; $j < ($this->smallBlockSize - $exp); $j++) {
                    $result .= "\x00";
                }
            }

            $item->setStartBlock($iSmBlk);
            $iSmBlk += $sbCount;
        }

        $sbCount = $this->getPointersPerBlock($this->bigBlockSize);
        if ($iSmBlk % $sbCount) {
            for ($i = 0; $i < ($sbCount - ($iSmBlk % $sbCount)); $i++) {
                $this->writeUlong(-1);
            }
        }

        return $result;
    }

    /**
     * Saves all the PPS's WKs
     *
     * @param PPS[] $raList Reference to an array with all PPS's
     */
    protected function savePps(&$raList)
    {
        // Save each PPS WK
        $raListCount = count($raList);
        for ($i = 0; $i < $raListCount; $i++) {
            $this->write($raList[$i]->getPpsWk());
        }

        // Adjust for Block
        $iCnt = count($raList);
        $iBCnt = $this->getPointersPerBlock($this->bigBlockSize, OLE::PPS_SIZE);
        if ($iCnt % $iBCnt) {
            for ($i = 0; $i < (($iBCnt - ($iCnt % $iBCnt)) * OLE::PPS_SIZE); $i++) {
                $this->write("\x00");
            }
        }
    }

    /**
     * @param $value
     */
    protected function writeUlong($value)
    {
        $this->write(pack("V", $value));
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
        $headerEntriesCount = $info["header_blockchain_list_entries"];
        $entriesCount = $info["blockchain_list_entries"];
        $entriesPerBlock = $info["entries_per_block"];

        $data = "";

        if ($numSbBlocks > 0) {
            for ($i = 0; $i < ($numSbBlocks - 1); $i++) {
                $data .= pack("V", $i + 1);
            }
            $data .= pack("V", -2);
        }

        for ($i = 0; $i < ($numBbBlocks - 1); $i++) {
            $data .= pack("V", $i + $numSbBlocks + 1);
        }
        $data .= pack("V", -2);

        for ($i = 0; $i < ($numPpsBlocks - 1); $i++) {
            $data .= pack("V", $i + $numSbBlocks + $numBbBlocks + 1);
        }
        $data .= pack("V", -2);

        for ($i = 0; $i < $info["FD_blockchain_entries"]; $i++) {
            $data .= pack("V", 0xFFFFFFFD);
        }

        for ($i = 0; $i < $info["FC_blockchain_entries"]; $i++) {
            $data .= pack("V", 0xFFFFFFFC);
        }

        // Adjust for Block
        $allEntries = $numSbBlocks + $numBbBlocks + $numPpsBlocks + $info["FD_blockchain_entries"]
            + $info["FC_blockchain_entries"];
        if ($allEntries % $entriesPerBlock) {
            $rest = $entriesPerBlock - ($allEntries % $entriesPerBlock);
            for ($i = 0; $i < $rest; $i++) {
                $data .= pack("V", -1);
            }
        }

        // Extra BDList
        if ($entriesCount > $headerEntriesCount) {
            $iN = 0;
            $iNb = 0;
            for ($i = $headerEntriesCount; $i < $entriesCount; $i++, $iN++) {
                if ($iN >= ($entriesPerBlock - 1)) {
                    $iN = 0;
                    $iNb++;
                    $data .= pack(
                        "V",
                        $numSbBlocks + $numBbBlocks + $numPpsBlocks + $info["FD_blockchain_entries"]
                        + $iNb
                    );
                }

                $data .= pack("V", $numBbBlocks + $numSbBlocks + $numPpsBlocks + $i);
            }

            $allEntries = $entriesCount - $headerEntriesCount;
            if (($allEntries % ($entriesPerBlock - 1))) {
                $rest = ($entriesPerBlock - 1) - ($allEntries % ($entriesPerBlock - 1));
                for ($i = 0; $i < $rest; $i++) {
                    $data .= pack("V", -1);
                }
            }

            $data .= pack("V", -2);
        }

        $this->write($data);
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
        $headerEntriesCount = $info["header_blockchain_list_entries"];
        $entriesCount = $info["blockchain_list_entries"];

        $this->write(
            "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
            . "\x00\x00\x00\x00"
            . "\x00\x00\x00\x00"
            . "\x00\x00\x00\x00"
            . "\x00\x00\x00\x00"
            . pack("v", 0x3b)
            . pack("v", 0x03)
            . pack("v", -2)
            . pack("v", 9)
            . pack("v", 6)
            . pack("v", 0)
            . "\x00\x00\x00\x00"
            . "\x00\x00\x00\x00"
            . pack("V", $entriesCount)
            . pack("V", $numSbBlocks + $numBbBlocks) //ROOT START
            . pack("V", 0)
            . pack("V", 0x1000)
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
            $this->writeUlong($numSbBlocks + $numBbBlocks + $numPpsBlocks + $info["FD_blockchain_entries"]);
            $this->writeUlong($info["FC_blockchain_entries"]);
        }

        // BDList
        for ($i = 0; $i < $headerEntriesCount && $i < $entriesCount; $i++) {
            $this->writeUlong($numBbBlocks + $numSbBlocks + $numPpsBlocks + $i);
        }

        if ($i < $headerEntriesCount) {
            for ($j = 0; $j < ($headerEntriesCount - $i); $j++) {
                $this->writeUlong(-1);
            }
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
            "header_blockchain_list_entries" => $this->getPointersPerBlock($this->bigBlockSize - 0x4C),
            "blockchain_entries" => $totalBlocks,
            "ext_blockchain_list_entries" => 0,
            "FC_blockchain_entries" => 0
        );

        $info["FD_blockchain_entries"] = $this->getNumberOfPointerBlocks(
            $info["blockchain_entries"]
        );
        $info["blockchain_list_entries"] = $this->getNumberOfPointerBlocks(
            $info["blockchain_entries"] + $info["FD_blockchain_entries"]
        );

        // do some magic
        if ($info["blockchain_list_entries"] <= $info["header_blockchain_list_entries"]) {
            return $info;
        }

        do {
            $info["blockchain_list_entries"] = $this->getNumberOfPointerBlocks(
                $info["blockchain_entries"] + $info["FD_blockchain_entries"]
                + $info["FC_blockchain_entries"]
            );
            $info["ext_blockchain_list_entries"]
                = $info["blockchain_list_entries"] - $info["header_blockchain_list_entries"];
            $info["FC_blockchain_entries"] = $this->getNumberOfPointerBlocks(
                $info["ext_blockchain_list_entries"]
            );
            $info["FD_blockchain_entries"] = $this->getNumberOfPointerBlocks(
                $totalBlocks + $info["FD_blockchain_entries"] + $info["FC_blockchain_entries"]
            );
        } while ($info["blockchain_list_entries"] < $this->getNumberOfPointerBlocks(
            $info["blockchain_entries"] + $info["FD_blockchain_entries"] + $info["FC_blockchain_entries"]
        ));

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
    protected function getPointersPerBlock($blockSize, $pointerSize = OLE::LONG_INT_SIZE)
    {
        return intval(floor($blockSize / $pointerSize));
    }
}
