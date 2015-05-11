<?php
/**
 * Created by PhpStorm.
 * User: mac
 * Date: 04.05.15
 * Time: 1:00
 */

namespace Xls;

class SharedStringsTable
{
    /**
     * Total number of strings
     * @var int
     */
    protected $totalCount = 0;

    /**
     * Number of unique strings
     * @var int
     */
    protected $uniqueCount = 0;

    /**
     * Array containing all the unique strings
     * @var array
     */
    protected $data = array();

    /**
     * @return int
     */
    public function getTotalCount()
    {
        return $this->totalCount;
    }

    /**
     * @return int
     */
    public function getUniqueCount()
    {
        return $this->uniqueCount;
    }

    /**
     * @return array
     */
    public function getStrings()
    {
        return array_keys($this->data);
    }

    /**
     * @param $str
     */
    public function add($str)
    {
        if (!isset($this->data[$str])) {
            $this->data[$str] = $this->uniqueCount++;
        }
        $this->totalCount++;
    }

    /**
     * @param $str
     *
     * @return mixed
     * @throws \Exception
     */
    public function getStrIdx($str)
    {
        if (isset($this->data[$str])) {
            return $this->data[$str];
        }

        throw new \Exception('String "'. $str . '" not found in Shared Strings Table');
    }

    /**
     * @param $str
     *
     * @return array
     */
    public function getStringInfo($str)
    {
        $info = unpack("vlength/Cunicode", $str);

        return array(
            'is_unicode' => $info["unicode"],
            'header_length' => ($info["unicode"] == 1) ? 4 : 3,
            'length' => strlen($str)
        );
    }

    /**
     * Handling of the SST continue blocks is complicated by the need to include an
     * additional continuation byte depending on whether the string is split between
     * blocks or whether it starts at the beginning of the block. (There are also
     * additional complications that will arise later when/if Rich Strings are
     * supported).
     *
     * @param null $tmpBlockSizes
     * @param bool $returnDataToWrite
     *
     * @return array
     */
    public function getBlocksSizesOrDataToWrite($tmpBlockSizes = null, $returnDataToWrite = false)
    {
        $continueLimit = Biff8::getContinueLimit();
        $blockLength = 0;
        $written = 0;
        $blockSizes = array();
        $data = array();
        $continue = 0;

        foreach ($this->getStrings() as $string) {
            $info = $this->getStringInfo($string);
            $splitString = 0;

            // Block length is the total length of the strings that will be
            // written out in a single SST or CONTINUE block.
            $blockLength += $info['length'];

            // We can write the string if it doesn't cross a CONTINUE boundary
            if ($blockLength < $continueLimit) {
                $data[] = $string;
                $written += $info['length'];
                continue;
            }

            // Deal with the cases where the next string to be written will exceed
            // the CONTINUE boundary. If the string is very long it may need to be
            // written in more than one CONTINUE record.
            while ($blockLength >= $continueLimit) {
                $spaceRemaining = $continueLimit - $written - $continue;

                // Unicode data should only be split on char (2 byte) boundaries.
                // Therefore, in some cases we need to reduce the amount of available
                // space by 1 byte to ensure the correct alignment.
                $align = 0;

                if ($spaceRemaining > $info['header_length']) {
                    // Only applies to Unicode strings
                    if ($info['is_unicode']) {
                        if (!$splitString && $spaceRemaining % 2 != 1) {
                            // String contains 3 byte header => split on odd boundary
                            $spaceRemaining--;
                            $align = 1;
                        } elseif ($splitString && $spaceRemaining % 2 == 1) {
                            // Split section without header => split on even boundary
                            $spaceRemaining--;
                            $align = 1;
                        }

                        $splitString = 1;
                    }

                    // Write as much as possible of the string in the current block
                    $data[] = substr($string, 0, $spaceRemaining);
                    $written += $spaceRemaining;

                    // The remainder will be written in the next block(s)
                    $string = substr($string, $spaceRemaining);

                    // Reduce the current block length by the amount written
                    $blockLength -= $continueLimit - $continue - $align;

                    // Store the max size for this block
                    $blockSizes[] = $continueLimit - $align;

                    // If the current string was split then the next CONTINUE block
                    // should have the string continue flag (grbit) set unless the
                    // split string fits exactly into the remaining space.
                    $continue = ($blockLength > 0) ? 1 : 0;
                } else {
                    // Store the max size for this block
                    $blockSizes[] = $written + $continue;

                    // Not enough space to start the string in the current block
                    $blockLength -= $continueLimit - $spaceRemaining - $continue;
                    $continue = 0;
                }

                // Write the CONTINUE block header
                if (!empty($tmpBlockSizes)) {
                    $length = array_shift($tmpBlockSizes);
                    $header = Record\ContinueRecord::getHeader($length);
                    if ($continue) {
                        $header .= pack('C', $info['is_unicode']);
                    }

                    $data[] = $header;
                }

                // If the string (or substr) is small enough we can write it in the
                // new CONTINUE block. Else, go through the loop again to write it in
                // one or more CONTINUE blocks
                if ($blockLength < $continueLimit) {
                    $data[] = $string;
                    $written = $blockLength;
                } else {
                    $written = 0;
                }
            }
        }

        // Store the max size for the last block unless it is empty
        if ($written + $continue) {
            $blockSizes[] = $written + $continue;
        }

        return ($returnDataToWrite) ? $data : $blockSizes;
    }

    /**
     * Calculate the total length of the SST and associated CONTINUEs (if any).
     * The SST record will have a length even if it contains no strings.
     * This length is required to set the offsets in the BOUNDSHEET records since
     * they must be written before the SST records
     *
     * @param $blockSizes
     *
     * @return int
     */
    public function calcSharedStringsTableLength($blockSizes)
    {
        $length = 12;

        if (!empty($blockSizes)) {
            $length += array_shift($blockSizes); // SST
        }

        while (!empty($blockSizes)) {
            $length += 4 + array_shift($blockSizes); // CONTINUEs
        }

        return $length;
    }

    /**
     * @param string $str
     * @param string $inputEncoding
     *
     * @return string
     */
    public function getPackedString($str, $inputEncoding)
    {
        if ($inputEncoding == 'UTF-16LE') {
            $strlen = function_exists('mb_strlen') ? mb_strlen($str, 'UTF-16LE') : (strlen($str) / 2);
            $encoding = 0x1;
        } elseif ($inputEncoding != '') {
            $str = iconv($inputEncoding, 'UTF-16LE', $str);
            $strlen = function_exists('mb_strlen') ? mb_strlen($str, 'UTF-16LE') : (strlen($str) / 2);
            $encoding = 0x1;
        } else {
            $strlen = strlen($str);
            $encoding = 0x0;
        }

        return pack('vC', $strlen, $encoding) . $str;
    }
}
