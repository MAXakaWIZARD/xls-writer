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
     * The temporary dir for storing the OLE file
     * @var string
     */
    public $tmpDir;

    /**
     * @var string
     */
    public $tmpFilename;

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
     * @access public
     * @param integer $time_1st A timestamp
     * @param integer $time_2nd A timestamp
     * @param array $raChild
     */
    public function __construct($time_1st, $time_2nd, $raChild)
    {
        $this->tmpDir = sys_get_temp_dir();

        parent::__construct(
            null,
            OLE::Asc2Ucs('Root Entry'),
            OLE_PPS_TYPE_ROOT,
            null,
            null,
            null,
            $time_1st,
            $time_2nd,
            null,
            $raChild
        );
    }

    /**
     * Sets the temp dir used for storing the OLE file
     *
     * @access public
     * @param string $dir The dir to be used as temp dir
     * @return true if given dir is valid, false otherwise
     */
    public function setTempDir($dir)
    {
        if (is_dir($dir)) {
            $this->tmpDir = $dir;
            return true;
        }
        return false;
    }

    /**
     * Method for saving the whole OLE container (including files).
     * In fact, if called with an empty argument (or '-'), it saves to a
     * temporary file and then outputs it's contents to stdout.
     *
     * @param string $filename The name of the file where to save the OLE container
     * @throws \Exception
     * @return mixed true on success
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
            if ($this->fileHandlerRoot == false) {
                throw new \Exception("Can't create temporary file.");
            }
        } else {
            $this->fileHandlerRoot = @fopen($filename, "wb");
            if ($this->fileHandlerRoot == false) {
                throw new \Exception("Can't open $filename. It may be in use or protected.");
            }
        }
        // Make an array of PPS's (for Save)
        $aList = array();
        self::savePpsSetPnt($aList, array($this));
        // calculate values for header
        list($iSBDcnt, $iBBcnt, $iPPScnt) = $this->calcSize($aList); //, $rhInfo);
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
     * @access private
     * @param OLE\PPS[] $raList Reference to an array of PPS's
     * @return array The array of numbers
     */
    public function calcSize(&$raList)
    {
        // Calculate Basic Setting
        list($iSBDcnt, $iBBcnt, $iPPScnt) = array(0, 0, 0);
        $iSmallLen = 0;
        $iSBcnt = 0;
        for ($i = 0; $i < count($raList); $i++) {
            if ($raList[$i]->Type == OLE_PPS_TYPE_FILE) {
                $raList[$i]->Size = $raList[$i]->dataLen();
                if ($raList[$i]->Size < OLE_DATA_SIZE_SMALL) {
                    $iSBcnt += floor($raList[$i]->Size / $this->smallBlockSize)
                        + (($raList[$i]->Size % $this->smallBlockSize) ? 1 : 0);
                } else {
                    $iBBcnt += (floor($raList[$i]->Size / $this->bigBlockSize) +
                        (($raList[$i]->Size % $this->bigBlockSize) ? 1 : 0));
                }
            }
        }
        $iSmallLen = $iSBcnt * $this->smallBlockSize;
        $iSlCnt = floor($this->bigBlockSize / OLE_LONG_INT_SIZE);
        $iSBDcnt = floor($iSBcnt / $iSlCnt) + (($iSBcnt % $iSlCnt) ? 1 : 0);
        $iBBcnt += (floor($iSmallLen / $this->bigBlockSize) +
            (($iSmallLen % $this->bigBlockSize) ? 1 : 0));
        $iCnt = count($raList);
        $iBdCnt = $this->bigBlockSize / OLE_PPS_SIZE;
        $iPPScnt = (floor($iCnt / $iBdCnt) + (($iCnt % $iBdCnt) ? 1 : 0));

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
            return $this->createHeader($iSBDcnt, $iBBcnt, $iPPScnt);
        }

        // Calculate Basic Setting
        $iBlCnt = $this->bigBlockSize / OLE_LONG_INT_SIZE;
        $i1stBdL = ($this->bigBlockSize - 0x4C) / OLE_LONG_INT_SIZE;

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
        $FILE = $this->fileHandlerRoot;

        // cycle through PPS's
        for ($i = 0; $i < count($raList); $i++) {
            if ($raList[$i]->Type != OLE_PPS_TYPE_DIR) {
                $raList[$i]->Size = $raList[$i]->dataLen();
                if (($raList[$i]->Size >= OLE_DATA_SIZE_SMALL)
                    || (($raList[$i]->Type == OLE_PPS_TYPE_ROOT)
                        && isset($raList[$i]->data))
                ) {
                    // Write Data
                    if (isset($raList[$i]->ppsFile)) {
                        $iLen = 0;
                        fseek($raList[$i]->ppsFile, 0); // To The Top
                        while ($sBuff = fread($raList[$i]->ppsFile, 4096)) {
                            $iLen += strlen($sBuff);
                            fwrite($FILE, $sBuff);
                        }
                    } else {
                        fwrite($FILE, $raList[$i]->data);
                    }

                    if ($raList[$i]->Size % $this->bigBlockSize) {
                        $loopEnd = ($this->bigBlockSize - ($raList[$i]->Size % $this->bigBlockSize));
                        for ($j = 0; $j < $loopEnd; $j++) {
                            fwrite($FILE, "\x00");
                        }
                    }
                    // Set For PPS
                    $raList[$i]->StartBlock = $iStBlk;
                    $iStBlk += (floor($raList[$i]->Size / $this->bigBlockSize) +
                        (($raList[$i]->Size % $this->bigBlockSize) ? 1 : 0));
                }
                // Close file for each PPS, and unlink it
                if (isset($raList[$i]->ppsFile)) {
                    @fclose($raList[$i]->ppsFile);
                    $raList[$i]->ppsFile = null;
                    @unlink($raList[$i]->tmpFilename);
                }
            }
        }
    }

    /**
     * get small data (PPS's with data smaller than OLE_DATA_SIZE_SMALL)
     *
     * @access private
     * @param OLE\PPS[] &$raList Reference to array of PPS's
     */
    public function makeSmallData(&$raList)
    {
        $sRes = '';
        $FILE = $this->fileHandlerRoot;
        $iSmBlk = 0;

        for ($i = 0; $i < count($raList); $i++) {
            // Make SBD, small data string
            if ($raList[$i]->Type == OLE_PPS_TYPE_FILE) {
                if ($raList[$i]->Size <= 0) {
                    continue;
                }

                if ($raList[$i]->Size < OLE_DATA_SIZE_SMALL) {
                    $iSmbCnt = floor($raList[$i]->Size / $this->smallBlockSize)
                        + (($raList[$i]->Size % $this->smallBlockSize) ? 1 : 0);
                    // Add to SBD
                    for ($j = 0; $j < ($iSmbCnt - 1); $j++) {
                        fwrite($FILE, pack("V", $j + $iSmBlk + 1));
                    }
                    fwrite($FILE, pack("V", -2));

                    // Add to Data String(this will be written for RootEntry)
                    if ($raList[$i]->ppsFile) {
                        fseek($raList[$i]->ppsFile, 0); // To The Top
                        while ($sBuff = fread($raList[$i]->ppsFile, 4096)) {
                            $sRes .= $sBuff;
                        }
                    } else {
                        $sRes .= $raList[$i]->data;
                    }

                    if ($raList[$i]->Size % $this->smallBlockSize) {
                        for (
                            $j = 0; $j < ($this->smallBlockSize - ($raList[$i]->Size % $this->smallBlockSize));
                            $j++
                        ) {
                            $sRes .= "\x00";
                        }
                    }
                    // Set for PPS
                    $raList[$i]->StartBlock = $iSmBlk;
                    $iSmBlk += $iSmbCnt;
                }
            }
        }

        $iSbCnt = floor($this->bigBlockSize / OLE_LONG_INT_SIZE);
        if ($iSmBlk % $iSbCnt) {
            for ($i = 0; $i < ($iSbCnt - ($iSmBlk % $iSbCnt)); $i++) {
                fwrite($FILE, pack("V", -1));
            }
        }

        return $sRes;
    }

    /**
     * Saves all the PPS's WKs
     *
     * @access private
     * @param OLE\PPS[] $raList Reference to an array with all PPS's
     */
    public function savePps(&$raList)
    {
        // Save each PPS WK
        for ($i = 0; $i < count($raList); $i++) {
            fwrite($this->fileHandlerRoot, $raList[$i]->getPpsWk());
        }
        // Adjust for Block
        $iCnt = count($raList);
        $iBCnt = $this->bigBlockSize / OLE_PPS_SIZE;
        if ($iCnt % $iBCnt) {
            for ($i = 0; $i < (($iBCnt - ($iCnt % $iBCnt)) * OLE_PPS_SIZE); $i++) {
                fwrite($this->fileHandlerRoot, "\x00");
            }
        }
    }

    /**
     * Saving Big Block Depot
     *
     * @access private
     * @param integer $iSbdSize
     * @param integer $iBsize
     * @param integer $iPpsCnt
     */
    public function saveBbd($iSbdSize, $iBsize, $iPpsCnt)
    {
        if ($this->newFunc) {
            return $this->createBigBlockChain($iSbdSize, $iBsize, $iPpsCnt);
        }

        $FILE = $this->fileHandlerRoot;
        // Calculate Basic Setting
        $iBbCnt = $this->bigBlockSize / OLE_LONG_INT_SIZE;
        $i1stBdL = ($this->bigBlockSize - 0x4C) / OLE_LONG_INT_SIZE;

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
                fwrite($FILE, pack("V", $i + 1));
            }
            fwrite($FILE, pack("V", -2));
        }
        // Set for B
        for ($i = 0; $i < ($iBsize - 1); $i++) {
            fwrite($FILE, pack("V", $i + $iSbdSize + 1));
        }
        fwrite($FILE, pack("V", -2));

        // Set for PPS
        for ($i = 0; $i < ($iPpsCnt - 1); $i++) {
            fwrite($FILE, pack("V", $i + $iSbdSize + $iBsize + 1));
        }
        fwrite($FILE, pack("V", -2));
        // Set for BBD itself ( 0xFFFFFFFD : BBD)
        for ($i = 0; $i < $iBdCnt; $i++) {
            fwrite($FILE, pack("V", 0xFFFFFFFD));
        }
        // Set for ExtraBDList
        for ($i = 0; $i < $iBdExL; $i++) {
            fwrite($FILE, pack("V", 0xFFFFFFFC));
        }
        // Adjust for Block
        if (($iAllW + $iBdCnt) % $iBbCnt) {
            for ($i = 0; $i < ($iBbCnt - (($iAllW + $iBdCnt) % $iBbCnt)); $i++) {
                fwrite($FILE, pack("V", -1));
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
                    fwrite($FILE, pack("V", $iAll + $iBdCnt + $iNb));
                }
                fwrite($FILE, pack("V", $iBsize + $iSbdSize + $iPpsCnt + $i));
            }
            if (($iBdCnt - $i1stBdL) % ($iBbCnt - 1)) {
                for ($i = 0; $i < (($iBbCnt - 1) - (($iBdCnt - $i1stBdL) % ($iBbCnt - 1))); $i++) {
                    fwrite($FILE, pack("V", -1));
                }
            }
            fwrite($FILE, pack("V", -2));
        }
    }

    /**
     * New method to store Bigblock chain
     *
     * @access private
     * @param integer $num_sb_blocks - number of Smallblock depot blocks
     * @param integer $num_bb_blocks - number of Bigblock depot blocks
     * @param integer $num_pps_blocks - number of PropertySetStorage blocks
     */
    public function createBigBlockChain($num_sb_blocks, $num_bb_blocks, $num_pps_blocks)
    {
        $bbd_info = $this->caclBigBlockChain($num_sb_blocks, $num_bb_blocks, $num_pps_blocks);

        $data = "";

        if ($num_sb_blocks > 0) {
            for ($i = 0; $i < ($num_sb_blocks - 1); $i++) {
                $data .= pack("V", $i + 1);
            }
            $data .= pack("V", -2);
        }

        for ($i = 0; $i < ($num_bb_blocks - 1); $i++) {
            $data .= pack("V", $i + $num_sb_blocks + 1);
        }
        $data .= pack("V", -2);

        for ($i = 0; $i < ($num_pps_blocks - 1); $i++) {
            $data .= pack("V", $i + $num_sb_blocks + $num_bb_blocks + 1);
        }
        $data .= pack("V", -2);

        for ($i = 0; $i < $bbd_info["0xFFFFFFFD_blockchain_entries"]; $i++) {
            $data .= pack("V", 0xFFFFFFFD);
        }

        for ($i = 0; $i < $bbd_info["0xFFFFFFFC_blockchain_entries"]; $i++) {
            $data .= pack("V", 0xFFFFFFFC);
        }

        // Adjust for Block
        $all_entries = $num_sb_blocks + $num_bb_blocks + $num_pps_blocks + $bbd_info["0xFFFFFFFD_blockchain_entries"]
            + $bbd_info["0xFFFFFFFC_blockchain_entries"];
        if ($all_entries % $bbd_info["entries_per_block"]) {
            $rest = $bbd_info["entries_per_block"] - ($all_entries % $bbd_info["entries_per_block"]);
            for ($i = 0; $i < $rest; $i++) {
                $data .= pack("V", -1);
            }
        }

        // Extra BDList
        if ($bbd_info["blockchain_list_entries"] > $bbd_info["header_blockchain_list_entries"]) {
            $iN = 0;
            $iNb = 0;
            for (
                $i = $bbd_info["header_blockchain_list_entries"]; $i < $bbd_info["blockchain_list_entries"]; $i++, $iN++
            ) {
                if ($iN >= ($bbd_info["entries_per_block"] - 1)) {
                    $iN = 0;
                    $iNb++;
                    $data .= pack(
                        "V",
                        $num_sb_blocks + $num_bb_blocks + $num_pps_blocks + $bbd_info["0xFFFFFFFD_blockchain_entries"]
                        + $iNb
                    );
                }

                $data .= pack("V", $num_bb_blocks + $num_sb_blocks + $num_pps_blocks + $i);
            }

            $all_entries = $bbd_info["blockchain_list_entries"] - $bbd_info["header_blockchain_list_entries"];
            if (($all_entries % ($bbd_info["entries_per_block"] - 1))) {
                $rest = ($bbd_info["entries_per_block"] - 1) - ($all_entries % ($bbd_info["entries_per_block"] - 1));
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
     * @access private
     * @param integer $num_sb_blocks - number of Smallblock depot blocks
     * @param integer $num_bb_blocks - number of Bigblock depot blocks
     * @param integer $num_pps_blocks - number of PropertySetStorage blocks
     */
    public function createHeader($num_sb_blocks, $num_bb_blocks, $num_pps_blocks)
    {
        $FILE = $this->fileHandlerRoot;

        $bbd_info = $this->caclBigBlockChain($num_sb_blocks, $num_bb_blocks, $num_pps_blocks);

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
            . pack("V", $bbd_info["blockchain_list_entries"])
            . pack("V", $num_sb_blocks + $num_bb_blocks) //ROOT START
            . pack("V", 0)
            . pack("V", 0x1000)
        );

        //Small Block Depot
        if ($num_sb_blocks > 0) {
            fwrite($FILE, pack("V", 0));
        } else {
            fwrite($FILE, pack("V", -2));
        }

        fwrite($FILE, pack("V", $num_sb_blocks));

        // Extra BDList Start, Count
        if ($bbd_info["blockchain_list_entries"] < $bbd_info["header_blockchain_list_entries"]) {
            fwrite(
                $FILE,
                pack("V", -2) . // Extra BDList Start
                pack("V", 0) // Extra BDList Count
            );
        } else {
            fwrite(
                $FILE,
                pack(
                    "V",
                    $num_sb_blocks + $num_bb_blocks + $num_pps_blocks + $bbd_info["0xFFFFFFFD_blockchain_entries"]
                ) . pack("V", $bbd_info["0xFFFFFFFC_blockchain_entries"])
            );
        }

        // BDList
        for (
            $i = 0; $i < $bbd_info["header_blockchain_list_entries"] and $i < $bbd_info["blockchain_list_entries"]; $i++
        ) {
            fwrite($FILE, pack("V", $num_bb_blocks + $num_sb_blocks + $num_pps_blocks + $i));
        }

        if ($i < $bbd_info["header_blockchain_list_entries"]) {
            for ($j = 0; $j < ($bbd_info["header_blockchain_list_entries"] - $i); $j++) {
                fwrite($FILE, (pack("V", -1)));
            }
        }
    }

    /**
     * New method to calculate Bigblock chain
     *
     * @access private
     * @param integer $num_sb_blocks - number of Smallblock depot blocks
     * @param integer $num_bb_blocks - number of Bigblock depot blocks
     * @param integer $num_pps_blocks - number of PropertySetStorage blocks
     */
    public function caclBigBlockChain($num_sb_blocks, $num_bb_blocks, $num_pps_blocks)
    {
        $bbd_info["entries_per_block"] = $this->bigBlockSize / OLE_LONG_INT_SIZE;
        $bbd_info["header_blockchain_list_entries"] = ($this->bigBlockSize - 0x4C) / OLE_LONG_INT_SIZE;
        $bbd_info["blockchain_entries"] = $num_sb_blocks + $num_bb_blocks + $num_pps_blocks;
        $bbd_info["0xFFFFFFFD_blockchain_entries"] = $this->getNumberOfPointerBlocks(
            $bbd_info["blockchain_entries"]
        );
        $bbd_info["blockchain_list_entries"] = $this->getNumberOfPointerBlocks(
            $bbd_info["blockchain_entries"] + $bbd_info["0xFFFFFFFD_blockchain_entries"]
        );

        // do some magic
        $bbd_info["ext_blockchain_list_entries"] = 0;
        $bbd_info["0xFFFFFFFC_blockchain_entries"] = 0;
        if ($bbd_info["blockchain_list_entries"] > $bbd_info["header_blockchain_list_entries"]) {
            do {
                $bbd_info["blockchain_list_entries"] = $this->getNumberOfPointerBlocks(
                    $bbd_info["blockchain_entries"] + $bbd_info["0xFFFFFFFD_blockchain_entries"]
                    + $bbd_info["0xFFFFFFFC_blockchain_entries"]
                );
                $bbd_info["ext_blockchain_list_entries"]
                    = $bbd_info["blockchain_list_entries"] - $bbd_info["header_blockchain_list_entries"];
                $bbd_info["0xFFFFFFFC_blockchain_entries"] = $this->getNumberOfPointerBlocks(
                    $bbd_info["ext_blockchain_list_entries"]
                );
                $bbd_info["0xFFFFFFFD_blockchain_entries"] = $this->getNumberOfPointerBlocks(
                    $num_sb_blocks + $num_bb_blocks + $num_pps_blocks + $bbd_info["0xFFFFFFFD_blockchain_entries"]
                    + $bbd_info["0xFFFFFFFC_blockchain_entries"]
                );
            } while ($bbd_info["blockchain_list_entries"] < $this->getNumberOfPointerBlocks(
                    $bbd_info["blockchain_entries"]
                    + $bbd_info["0xFFFFFFFD_blockchain_entries"]
                    + $bbd_info["0xFFFFFFFC_blockchain_entries"]
                )
            );
        }

        return $bbd_info;
    }

    /**
     * Calculates number of pointer blocks
     *
     * @param integer $num_pointers - number of pointers
     * @return int
     */
    public function getNumberOfPointerBlocks($num_pointers)
    {
        $pointers_per_block = $this->bigBlockSize / OLE_LONG_INT_SIZE;

        return floor($num_pointers / $pointers_per_block) + (($num_pointers % $pointers_per_block) ? 1 : 0);
    }

    /**
     * Support method for some hexdumping
     *
     * @access public
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

            $chars = array();
        }
    }
}
