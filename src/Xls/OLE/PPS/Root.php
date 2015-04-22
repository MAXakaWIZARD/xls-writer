<?php

namespace Xls\OLE\PPS;

use Xls\OLE;

/**
 * Class for creating Root PPS's for OLE containers
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 * @category Structures
 * @package  OLE
 */
class Root extends OLE\PPS
{
    /**
     * Flag to enable new logic
     * @var bool
     */
    public $newFunc = true;

    /**
     * @var
     */
    public $smallBlockSize;

    /**
     * @var
     */
    public $bigBlockSize;

    /**
     * @var
     */
    public $fileHandlerRoot;

    /**
     * Constructor
     *
     * @param integer $time1st A timestamp
     * @param integer $time2nd A timestamp
     * @param File[] $raChild
     */
    public function __construct(
        $time1st = null,
        $time2nd = null,
        $raChild = array()
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
            $raChild
        );
    }

    /**
     * Method for saving the whole OLE container (including files).
     * In fact, if called with an empty argument (or '-'), it saves to a
     * temporary file and then outputs it's contents to stdout.
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

        // Open temp file if we are sending output to stdout
        if (($filename == '-') || ($filename == '')) {
            $this->tmpFilename = tempnam($this->tmpDir, "OLE_PPS_Root");
            $this->fileHandlerRoot = @fopen($this->tmpFilename, "w+b");
            if ($this->fileHandlerRoot === false) {
                throw new \Exception("Can't create temporary file.");
            }
        } else {
            $this->fileHandlerRoot = @fopen($filename, "wb");
            if ($this->fileHandlerRoot === false) {
                throw new \Exception("Can't open $filename. It may be in use or protected.");
            }
        }

        // Make an array of PPS's (for Save)
        $aList = array();
        self::savePpsSetPnt($aList, array($this));
        // calculate values for header
        list($iSBDcnt, $iBBcnt, $iPPScnt) = $this->calcSize($aList);
        // Save Header
        $this->saveHeader($iSBDcnt, $iBBcnt, $iPPScnt);

        // Make Small Data string (write SBD)
        $this->data = $this->makeSmallData($aList);

        // Write BB
        $this->saveBigData($iSBDcnt, $aList);
        // Write PPS
        $this->savePps($aList);
        // Write Big Block Depot and BDList and Adding Header informations
        $this->saveBbd($iSBDcnt, $iBBcnt, $iPPScnt);
        // Close File, send it to stdout if necessary
        if (($filename == '-') || ($filename == '')) {
            fseek($this->fileHandlerRoot, 0);
            fpassthru($this->fileHandlerRoot);
            @fclose($this->fileHandlerRoot);
            // Delete the temporary file.
            @unlink($this->tmpFilename);
        } else {
            @fclose($this->fileHandlerRoot);
        }

        return true;
    }

    /**
     * Calculate some numbers
     *
     * @param OLE\PPS[] $list Reference to an array of PPS's
     *
     * @return array The array of numbers
     */
    protected function calcSize($list)
    {
        $iSBcnt = 0;
        $iBBcnt = 0;
        foreach ($list as $item) {
            if ($item->Type == OLE::PPS_TYPE_FILE) {
                $item->Size = $item->dataLen();
                if ($item->Size < OLE::DATA_SIZE_SMALL) {
                    $iSBcnt += floor($item->Size / $this->smallBlockSize)
                        + (($item->Size % $this->smallBlockSize) ? 1 : 0);
                } else {
                    $iBBcnt += (floor($item->Size / $this->bigBlockSize) +
                        (($item->Size % $this->bigBlockSize) ? 1 : 0));
                }
            }
        }

        $iSmallLen = $iSBcnt * $this->smallBlockSize;
        $iSlCnt = floor($this->bigBlockSize / OLE::LONG_INT_SIZE);
        $iSBDcnt = floor($iSBcnt / $iSlCnt) + (($iSBcnt % $iSlCnt) ? 1 : 0);
        $iBBcnt += floor($iSmallLen / $this->bigBlockSize) +
            (($iSmallLen % $this->bigBlockSize) ? 1 : 0);
        $iCnt = count($list);
        $iBdCnt = $this->bigBlockSize / OLE::PPS_SIZE;
        $iPPScnt = floor($iCnt / $iBdCnt) + (($iCnt % $iBdCnt) ? 1 : 0);

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
        $FILE = $this->fileHandlerRoot;

        if ($this->newFunc) {
            $this->createHeader($iSBDcnt, $iBBcnt, $iPPScnt);
            return;
        }

        // Calculate Basic Setting
        $iBlCnt = $this->bigBlockSize / OLE::LONG_INT_SIZE;
        $i1stBdL = ($this->bigBlockSize - 0x4C) / OLE::LONG_INT_SIZE;

        $iBdExL = 0;
        $iAll = $iBBcnt + $iPPScnt + $iSBDcnt;
        $iAllW = $iAll;
        $iBdCntW = floor($iAllW / $iBlCnt) + (($iAllW % $iBlCnt) ? 1 : 0);
        $iBdCnt = floor(($iAll + $iBdCntW) / $iBlCnt) + ((($iAllW + $iBdCntW) % $iBlCnt) ? 1 : 0);

        // Calculate BD count
        if ($iBdCnt > $i1stBdL) {
            while (1) {
                $iBdExL++;
                $iAllW++;
                $iBdCntW = floor($iAllW / $iBlCnt) + (($iAllW % $iBlCnt) ? 1 : 0);
                $iBdCnt = floor(($iAllW + $iBdCntW) / $iBlCnt) + ((($iAllW + $iBdCntW) % $iBlCnt) ? 1 : 0);
                if ($iBdCnt <= ($iBdExL * $iBlCnt + $i1stBdL)) {
                    break;
                }
            }
        }

        // Save Header
        fwrite(
            $FILE,
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
            fwrite(
                $FILE,
                pack("V", -2) . // Extra BDList Start
                pack("V", 0) // Extra BDList Count
            );
        } else {
            fwrite($FILE, pack("V", $iAll + $iBdCnt) . pack("V", $iBdExL));
        }

        // BDList
        for ($i = 0; $i < $i1stBdL && $i < $iBdCnt; $i++) {
            fwrite($FILE, pack("V", $iAll + $i));
        }

        if ($i < $i1stBdL) {
            for ($j = 0; $j < ($i1stBdL - $i); $j++) {
                fwrite($FILE, (pack("V", -1)));
            }
        }
    }

    /**
     * Saving big data (PPS's with data bigger than OLE_DATA_SIZE_SMALL)
     *
     * @param integer $iStBlk
     * @param OLE\PPS[] &$raList Reference to array of PPS's
     */
    public function saveBigData($iStBlk, &$raList)
    {
        // loop through PPS's
        $raListCount = count($raList);
        for ($i = 0; $i < $raListCount; $i++) {
            if ($raList[$i]->Type != OLE::PPS_TYPE_DIR) {
                $raList[$i]->Size = $raList[$i]->dataLen();
                if (($raList[$i]->Size >= OLE::DATA_SIZE_SMALL)
                    || (($raList[$i]->Type == OLE::PPS_TYPE_ROOT)
                        && isset($raList[$i]->data))
                ) {
                    // Write Data
                    $ppsFile = $raList[$i]->getPpsFile();
                    if (is_resource($ppsFile)) {
                        $iLen = 0;
                        fseek($ppsFile, 0); // To The Top
                        while ($sBuff = fread($ppsFile, 4096)) {
                            $iLen += strlen($sBuff);
                            fwrite($this->fileHandlerRoot, $sBuff);
                        }
                    } else {
                        fwrite($this->fileHandlerRoot, $raList[$i]->data);
                    }

                    if ($raList[$i]->Size % $this->bigBlockSize) {
                        $loopEnd = ($this->bigBlockSize - ($raList[$i]->Size % $this->bigBlockSize));
                        for ($j = 0; $j < $loopEnd; $j++) {
                            fwrite($this->fileHandlerRoot, "\x00");
                        }
                    }
                    // Set For PPS
                    $raList[$i]->StartBlock = $iStBlk;
                    $iStBlk += (floor($raList[$i]->Size / $this->bigBlockSize) +
                        (($raList[$i]->Size % $this->bigBlockSize) ? 1 : 0));
                }
                $raList[$i]->removeTmpFile();
            }
        }
    }

    /**
     * get small data (PPS's with data smaller than OLE_DATA_SIZE_SMALL)
     *
     * @param OLE\PPS[] &$raList Reference to array of PPS's
     * @return string
     */
    protected function makeSmallData(&$raList)
    {
        $sRes = '';
        $file = $this->fileHandlerRoot;
        $iSmBlk = 0;
        $raListCount = count($raList);
        for ($i = 0; $i < $raListCount; $i++) {
            // Make SBD, small data string
            if ($raList[$i]->Type == OLE::PPS_TYPE_FILE) {
                if ($raList[$i]->Size <= 0) {
                    continue;
                }

                if ($raList[$i]->Size < OLE::DATA_SIZE_SMALL) {
                    $iSmbCnt = floor($raList[$i]->Size / $this->smallBlockSize)
                        + (($raList[$i]->Size % $this->smallBlockSize) ? 1 : 0);
                    // Add to SBD
                    for ($j = 0; $j < ($iSmbCnt - 1); $j++) {
                        fwrite($file, pack("V", $j + $iSmBlk + 1));
                    }
                    fwrite($file, pack("V", -2));

                    // Add to Data String(this will be written for RootEntry)
                    if ($raList[$i]->ppsFile) {
                        fseek($raList[$i]->ppsFile, 0); // To The Top
                        while ($sBuff = fread($raList[$i]->ppsFile, 4096)) {
                            $sRes .= $sBuff;
                        }
                    } else {
                        $sRes .= $raList[$i]->data;
                    }

                    $exp = $raList[$i]->Size % $this->smallBlockSize;
                    if ($exp) {
                        for ($j = 0; $j < ($this->smallBlockSize - $exp); $j++) {
                            $sRes .= "\x00";
                        }
                    }
                    // Set for PPS
                    $raList[$i]->StartBlock = $iSmBlk;
                    $iSmBlk += $iSmbCnt;
                }
            }
        }

        $iSbCnt = floor($this->bigBlockSize / OLE::LONG_INT_SIZE);
        if ($iSmBlk % $iSbCnt) {
            for ($i = 0; $i < ($iSbCnt - ($iSmBlk % $iSbCnt)); $i++) {
                fwrite($file, pack("V", -1));
            }
        }

        return $sRes;
    }

    /**
     * Saves all the PPS's WKs
     *
     * @param OLE\PPS[] $raList Reference to an array with all PPS's
     */
    protected function savePps(&$raList)
    {
        // Save each PPS WK
        $raListCount = count($raList);
        for ($i = 0; $i < $raListCount; $i++) {
            fwrite($this->fileHandlerRoot, $raList[$i]->getPpsWk());
        }
        // Adjust for Block
        $iCnt = count($raList);
        $iBCnt = $this->bigBlockSize / OLE::PPS_SIZE;
        if ($iCnt % $iBCnt) {
            for ($i = 0; $i < (($iBCnt - ($iCnt % $iBCnt)) * OLE::PPS_SIZE); $i++) {
                fwrite($this->fileHandlerRoot, "\x00");
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

        $file = $this->fileHandlerRoot;

        // Calculate Basic Setting
        $iBbCnt = $this->bigBlockSize / OLE::LONG_INT_SIZE;
        $i1stBdL = ($this->bigBlockSize - 0x4C) / OLE::LONG_INT_SIZE;

        $iBdExL = 0;
        $iAll = $iBsize + $iPpsCnt + $iSbdSize;
        $iAllW = $iAll;
        $iBdCntW = floor($iAllW / $iBbCnt) + (($iAllW % $iBbCnt) ? 1 : 0);
        $iBdCnt = floor(($iAll + $iBdCntW) / $iBbCnt) + ((($iAllW + $iBdCntW) % $iBbCnt) ? 1 : 0);
        // Calculate BD count
        if ($iBdCnt > $i1stBdL) {
            while (1) {
                $iBdExL++;
                $iAllW++;
                $iBdCntW = floor($iAllW / $iBbCnt) + (($iAllW % $iBbCnt) ? 1 : 0);
                $iBdCnt = floor(($iAllW + $iBdCntW) / $iBbCnt) + ((($iAllW + $iBdCntW) % $iBbCnt) ? 1 : 0);
                if ($iBdCnt <= ($iBdExL * $iBbCnt + $i1stBdL)) {
                    break;
                }
            }
        }

        // Making BD
        // Set for SBD
        if ($iSbdSize > 0) {
            for ($i = 0; $i < ($iSbdSize - 1); $i++) {
                fwrite($file, pack("V", $i + 1));
            }
            fwrite($file, pack("V", -2));
        }
        // Set for B
        for ($i = 0; $i < ($iBsize - 1); $i++) {
            fwrite($file, pack("V", $i + $iSbdSize + 1));
        }
        fwrite($file, pack("V", -2));

        // Set for PPS
        for ($i = 0; $i < ($iPpsCnt - 1); $i++) {
            fwrite($file, pack("V", $i + $iSbdSize + $iBsize + 1));
        }
        fwrite($file, pack("V", -2));
        // Set for BBD itself ( 0xFFFFFFFD : BBD)
        for ($i = 0; $i < $iBdCnt; $i++) {
            fwrite($file, pack("V", 0xFFFFFFFD));
        }
        // Set for ExtraBDList
        for ($i = 0; $i < $iBdExL; $i++) {
            fwrite($file, pack("V", 0xFFFFFFFC));
        }
        // Adjust for Block
        if (($iAllW + $iBdCnt) % $iBbCnt) {
            for ($i = 0; $i < ($iBbCnt - (($iAllW + $iBdCnt) % $iBbCnt)); $i++) {
                fwrite($file, pack("V", -1));
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
                    fwrite($file, pack("V", $iAll + $iBdCnt + $iNb));
                }
                fwrite($file, pack("V", $iBsize + $iSbdSize + $iPpsCnt + $i));
            }
            if (($iBdCnt - $i1stBdL) % ($iBbCnt - 1)) {
                for ($i = 0; $i < (($iBbCnt - 1) - (($iBdCnt - $i1stBdL) % ($iBbCnt - 1))); $i++) {
                    fwrite($file, pack("V", -1));
                }
            }
            fwrite($file, pack("V", -2));
        }
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
        $bbdInfo = $this->caclBigBlockChain($numSbBlocks, $numBbBlocks, $numPpsBlocks);

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

        for ($i = 0; $i < $bbdInfo["0xFFFFFFFD_blockchain_entries"]; $i++) {
            $data .= pack("V", 0xFFFFFFFD);
        }

        for ($i = 0; $i < $bbdInfo["0xFFFFFFFC_blockchain_entries"]; $i++) {
            $data .= pack("V", 0xFFFFFFFC);
        }

        // Adjust for Block
        $allEntries = $numSbBlocks + $numBbBlocks + $numPpsBlocks + $bbdInfo["0xFFFFFFFD_blockchain_entries"]
            + $bbdInfo["0xFFFFFFFC_blockchain_entries"];
        if ($allEntries % $bbdInfo["entries_per_block"]) {
            $rest = $bbdInfo["entries_per_block"] - ($allEntries % $bbdInfo["entries_per_block"]);
            for ($i = 0; $i < $rest; $i++) {
                $data .= pack("V", -1);
            }
        }

        // Extra BDList
        if ($bbdInfo["blockchain_list_entries"] > $bbdInfo["header_blockchain_list_entries"]) {
            $iN = 0;
            $iNb = 0;
            for (
                $i = $bbdInfo["header_blockchain_list_entries"]; $i < $bbdInfo["blockchain_list_entries"]; $i++, $iN++
            ) {
                if ($iN >= ($bbdInfo["entries_per_block"] - 1)) {
                    $iN = 0;
                    $iNb++;
                    $data .= pack(
                        "V",
                        $numSbBlocks + $numBbBlocks + $numPpsBlocks + $bbdInfo["0xFFFFFFFD_blockchain_entries"]
                        + $iNb
                    );
                }

                $data .= pack("V", $numBbBlocks + $numSbBlocks + $numPpsBlocks + $i);
            }

            $allEntries = $bbdInfo["blockchain_list_entries"] - $bbdInfo["header_blockchain_list_entries"];
            if (($allEntries % ($bbdInfo["entries_per_block"] - 1))) {
                $rest = ($bbdInfo["entries_per_block"] - 1) - ($allEntries % ($bbdInfo["entries_per_block"] - 1));
                for ($i = 0; $i < $rest; $i++) {
                    $data .= pack("V", -1);
                }
            }

            $data .= pack("V", -2);
        }

        /*
          $this->dump($data, 0, strlen($data));
          die;
        */

        fwrite($this->fileHandlerRoot, $data);
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
        $file = $this->fileHandlerRoot;

        $bbdInfo = $this->caclBigBlockChain($numSbBlocks, $numBbBlocks, $numPpsBlocks);

        // Save Header
        fwrite(
            $file,
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
            . pack("V", $bbdInfo["blockchain_list_entries"])
            . pack("V", $numSbBlocks + $numBbBlocks) //ROOT START
            . pack("V", 0)
            . pack("V", 0x1000)
        );

        //Small Block Depot
        if ($numSbBlocks > 0) {
            fwrite($file, pack("V", 0));
        } else {
            fwrite($file, pack("V", -2));
        }

        fwrite($file, pack("V", $numSbBlocks));

        // Extra BDList Start, Count
        if ($bbdInfo["blockchain_list_entries"] < $bbdInfo["header_blockchain_list_entries"]) {
            fwrite(
                $file,
                pack("V", -2) . // Extra BDList Start
                pack("V", 0) // Extra BDList Count
            );
        } else {
            fwrite(
                $file,
                pack(
                    "V",
                    $numSbBlocks + $numBbBlocks + $numPpsBlocks + $bbdInfo["0xFFFFFFFD_blockchain_entries"]
                ) . pack("V", $bbdInfo["0xFFFFFFFC_blockchain_entries"])
            );
        }

        // BDList
        for (
            $i = 0; $i < $bbdInfo["header_blockchain_list_entries"] && $i < $bbdInfo["blockchain_list_entries"]; $i++
        ) {
            fwrite($file, pack("V", $numBbBlocks + $numSbBlocks + $numPpsBlocks + $i));
        }

        if ($i < $bbdInfo["header_blockchain_list_entries"]) {
            for ($j = 0; $j < ($bbdInfo["header_blockchain_list_entries"] - $i); $j++) {
                fwrite($file, (pack("V", -1)));
            }
        }
    }

    /**
     * New method to calculate Bigblock chain
     *
     * @param integer $numSb - number of Smallblock depot blocks
     * @param integer $numBb - number of Bigblock depot blocks
     * @param integer $numPps - number of PropertySetStorage blocks
     */
    protected function caclBigBlockChain($numSb, $numBb, $numPps)
    {
        $bbdInfo["entries_per_block"] = $this->bigBlockSize / OLE::LONG_INT_SIZE;
        $bbdInfo["header_blockchain_list_entries"] = ($this->bigBlockSize - 0x4C) / OLE::LONG_INT_SIZE;
        $bbdInfo["blockchain_entries"] = $numSb + $numBb + $numPps;
        $bbdInfo["0xFFFFFFFD_blockchain_entries"] = $this->getNumberOfPointerBlocks(
            $bbdInfo["blockchain_entries"]
        );
        $bbdInfo["blockchain_list_entries"] = $this->getNumberOfPointerBlocks(
            $bbdInfo["blockchain_entries"] + $bbdInfo["0xFFFFFFFD_blockchain_entries"]
        );

        // do some magic
        $bbdInfo["ext_blockchain_list_entries"] = 0;
        $bbdInfo["0xFFFFFFFC_blockchain_entries"] = 0;
        if ($bbdInfo["blockchain_list_entries"] > $bbdInfo["header_blockchain_list_entries"]) {
            do {
                $bbdInfo["blockchain_list_entries"] = $this->getNumberOfPointerBlocks(
                    $bbdInfo["blockchain_entries"] + $bbdInfo["0xFFFFFFFD_blockchain_entries"]
                    + $bbdInfo["0xFFFFFFFC_blockchain_entries"]
                );
                $bbdInfo["ext_blockchain_list_entries"]
                    = $bbdInfo["blockchain_list_entries"] - $bbdInfo["header_blockchain_list_entries"];
                $bbdInfo["0xFFFFFFFC_blockchain_entries"] = $this->getNumberOfPointerBlocks(
                    $bbdInfo["ext_blockchain_list_entries"]
                );
                $bbdInfo["0xFFFFFFFD_blockchain_entries"] = $this->getNumberOfPointerBlocks(
                    $numSb + $numBb + $numPps + $bbdInfo["0xFFFFFFFD_blockchain_entries"]
                    + $bbdInfo["0xFFFFFFFC_blockchain_entries"]
                );
            } while ($bbdInfo["blockchain_list_entries"] < $this->getNumberOfPointerBlocks(
                    $bbdInfo["blockchain_entries"]
                    + $bbdInfo["0xFFFFFFFD_blockchain_entries"]
                    + $bbdInfo["0xFFFFFFFC_blockchain_entries"]
                )
            );
        }

        return $bbdInfo;
    }

    /**
     * Calculates number of pointer blocks
     *
     * @param integer $numPointers - number of pointers
     *
*@return int
     */
    public function getNumberOfPointerBlocks($numPointers)
    {
        $pointersPerBlock = $this->bigBlockSize / OLE::LONG_INT_SIZE;

        return floor($numPointers / $pointersPerBlock) + (($numPointers % $pointersPerBlock) ? 1 : 0);
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
