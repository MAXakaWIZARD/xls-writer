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
     * Flag to enable new logic
     * @var bool
     */
    protected $newFunc = true;

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
     * Constructor
     *
     * @param integer $time1st A timestamp
     * @param integer $time2nd A timestamp
     * @param PpsFile[] $children
     */
    public function __construct(
        $time1st = null,
        $time2nd = null,
        $children = array()
    ) {
        parent::__construct(
            null,
            OLE::asc2Ucs('Root Entry'),
            OLE::PPS_TYPE_ROOT,
            null,
            null,
            null,
            $time1st,
            $time2nd,
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
        // Initial Setting for saving
        $this->bigBlockSize = pow(
            2,
            ((isset($this->bigBlockSize)) ? $this->adjust2($this->bigBlockSize) : 9)
        );
        $this->smallBlockSize = pow(
            2,
            ((isset($this->smallBlockSize)) ? $this->adjust2($this->smallBlockSize) : 6)
        );

        $this->rootFilePointer = @fopen($filename, "wb");
        if ($this->rootFilePointer === false) {
            throw new \Exception("Can't open $filename. It may be in use or protected.");
        }

        // Make an array of PPS's (for Save)
        $aList = array();
        self::savePpsSetPnt($aList, array($this));
        // calculate values for header
        list($iSBDcnt, $iBBcnt, $iPPScnt) = $this->calcSize($aList);
        // Save Header
        $this->saveHeader($iSBDcnt, $iBBcnt, $iPPScnt);

        // Make Small Data string (write SBD)
        $this->data = $this->getAndWriteSmallData($aList);

        $this->saveBigData($iSBDcnt, $aList);
        $this->savePps($aList);
        $this->saveBbd($iSBDcnt, $iBBcnt, $iPPScnt);

        fclose($this->rootFilePointer);

        return true;
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
     * Helper function for caculating a magic value for block sizes
     *
     * @param integer $i2 The argument
     * @see save()
     * @return integer
     */
    public function adjust2($i2)
    {
        $iWk = log($i2) / log(2);

        return ($iWk > floor($iWk)) ? floor($iWk) + 1 : $iWk;
    }

    /**
     * Save OLE header
     *
     * @param integer $iSBDcnt
     * @param integer $iBBcnt
     * @param integer $iPPScnt
     */
    public function saveHeader($iSBDcnt, $iBBcnt, $iPPScnt)
    {
        if ($this->newFunc) {
            $this->createHeader($iSBDcnt, $iBBcnt, $iPPScnt);
            return;
        }

        // Calculate Basic Setting
        $iBlCnt = $this->getPointersPerBlock($this->bigBlockSize);
        $i1stBdL = $this->getPointersPerBlock($this->bigBlockSize - 0x4C);

        $iBdExL = 0;
        $iAll = $iBBcnt + $iPPScnt + $iSBDcnt;
        $iAllW = $iAll;
        $iBdCntW = $this->getBlocksCount($iAllW, $iBlCnt);
        $iBdCnt = $this->getBlocksCount($iAll + $iBdCntW, $iBlCnt);

        // Calculate BD count
        if ($iBdCnt > $i1stBdL) {
            while (true) {
                $iBdExL++;
                $iAllW++;
                $iBdCntW = $this->getBlocksCount($iAllW, $iBlCnt);
                $iBdCnt = $this->getBlocksCount($iAll + $iBdCntW, $iBlCnt);
                if ($iBdCnt <= ($iBdExL * $iBlCnt + $i1stBdL)) {
                    break;
                }
            }
        }

        // Save Header
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
            . pack("V", $iBdCnt)
            . pack("V", $iBBcnt + $iSBDcnt) //ROOT START
            . pack("V", 0)
            . pack("V", 0x1000)
            . pack("V", $iSBDcnt ? 0 : -2) //Small Block Depot
            . pack("V", $iSBDcnt)
        );

        // Extra BDList Start, Count
        if ($iBdCnt < $i1stBdL) {
            $this->writeUlong(-2); // Extra BDList Start
            $this->writeUlong(0); // Extra BDList Count
        } else {
            $this->writeUlong($iAll + $iBdCnt);
            $this->writeUlong($iBdExL);
        }

        // BDList
        for ($i = 0; $i < $i1stBdL && $i < $iBdCnt; $i++) {
            $this->writeUlong($iAll + $i);
        }

        if ($i < $i1stBdL) {
            for ($j = 0; $j < ($i1stBdL - $i); $j++) {
                $this->writeUlong(-1);
            }
        }
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
     * Saving Big Block Depot
     *
     * @param integer $iSbdSize
     * @param integer $iBsize
     * @param integer $iPpsCnt
     */
    protected function saveBbd($iSbdSize, $iBsize, $iPpsCnt)
    {
        if ($this->newFunc) {
            $this->createBigBlockChain($iSbdSize, $iBsize, $iPpsCnt);
            return;
        }

        // Calculate Basic Setting
        $iBbCnt = $this->getPointersPerBlock($this->bigBlockSize);
        $i1stBdL = $this->getPointersPerBlock($this->bigBlockSize - 0x4C);

        $iBdExL = 0;
        $iAll = $iBsize + $iPpsCnt + $iSbdSize;
        $iAllW = $iAll;
        $iBdCntW = $this->getBlocksCount($iAllW, $iBbCnt);
        $iBdCnt = $this->getBlocksCount($iAllW + $iBdCntW, $iBbCnt);
        // Calculate BD count
        if ($iBdCnt > $i1stBdL) {
            while (1) {
                $iBdExL++;
                $iAllW++;
                $iBdCntW = $this->getBlocksCount($iAllW, $iBbCnt);
                $iBdCnt = $this->getBlocksCount($iAllW + $iBdCntW, $iBbCnt);
                if ($iBdCnt <= ($iBdExL * $iBbCnt + $i1stBdL)) {
                    break;
                }
            }
        }

        // Making BD
        // Set for SBD
        if ($iSbdSize > 0) {
            for ($i = 0; $i < ($iSbdSize - 1); $i++) {
                $this->writeUlong($i + 1);
            }
            $this->writeUlong(-2);
        }

        // Set for B
        for ($i = 0; $i < ($iBsize - 1); $i++) {
            $this->writeUlong($i + $iSbdSize + 1);
        }
        $this->writeUlong(-2);

        // Set for PPS
        for ($i = 0; $i < ($iPpsCnt - 1); $i++) {
            $this->writeUlong($i + $iSbdSize + $iBsize + 1);
        }
        $this->writeUlong(-2);

        // Set for BBD itself ( 0xFFFFFFFD : BBD)
        for ($i = 0; $i < $iBdCnt; $i++) {
            $this->writeUlong(0xFFFFFFFD);
        }

        // Set for ExtraBDList
        for ($i = 0; $i < $iBdExL; $i++) {
            $this->writeUlong(0xFFFFFFFC);
        }

        // Adjust for Block
        if (($iAllW + $iBdCnt) % $iBbCnt) {
            for ($i = 0; $i < ($iBbCnt - (($iAllW + $iBdCnt) % $iBbCnt)); $i++) {
                $this->writeUlong(-1);
            }
        }

        // Extra BDList
        if ($iBdCnt > $i1stBdL) {
            $iN = 0;
            $iNb = 0;
            for ($i = $i1stBdL; $i < $iBdCnt; $i++, $iN++) {
                if ($iN >= ($iBbCnt - 1)) {
                    $iN = 0;
                    $iNb++;
                    $this->writeUlong($iAll + $iBdCnt + $iNb);
                }
                $this->writeUlong($iBsize + $iSbdSize + $iPpsCnt + $i);
            }
            if (($iBdCnt - $i1stBdL) % ($iBbCnt - 1)) {
                for ($i = 0; $i < (($iBbCnt - 1) - (($iBdCnt - $i1stBdL) % ($iBbCnt - 1))); $i++) {
                    $this->writeUlong(-1);
                }
            }
            $this->writeUlong(-2);
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
     * New method to store Bigblock chain
     *
     * @param integer $numSbBlocks - number of Smallblock depot blocks
     * @param integer $numBbBlocks - number of Bigblock depot blocks
     * @param integer $numPpsBlocks - number of PropertySetStorage blocks
     */
    protected function createBigBlockChain($numSbBlocks, $numBbBlocks, $numPpsBlocks)
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
     * New method to store Header
     *
     * @param integer $numSbBlocks - number of Smallblock depot blocks
     * @param integer $numBbBlocks - number of Bigblock depot blocks
     * @param integer $numPpsBlocks - number of PropertySetStorage blocks
     */
    public function createHeader($numSbBlocks, $numBbBlocks, $numPpsBlocks)
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

    /**
     * Support method for some hexdumping
     *
     * @param string $data - Binary data
     * @param integer $from - Start offset of data to dump
     * @param integer $to - Target offset of data to dump
     */
    public function dump($data, $from, $to)
    {
        $chars = array();
        for ($i = $from; $i < $to; $i++) {
            if (sizeof($chars) == 16) {
                printf("%08X (% 12d) |", $i - 16, $i - 16);
                foreach ($chars as $char) {
                    printf(" %02X", $char);
                }
                print " |\n";

                $chars = array();
            }

            $chars[] = ord($data[$i]);
        }

        if (sizeof($chars)) {
            printf("%08X (% 12d) |", $i - sizeof($chars), $i - sizeof($chars));
            foreach ($chars as $char) {
                printf(" %02X", $char);
            }
            print " |\n";
        }
    }
}
