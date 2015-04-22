<?php

namespace Xls\Writer;

/**
 * Class for writing Excel BIFF records.
 *
 * From "MICROSOFT EXCEL BINARY FILE FORMAT" by Mark O'Brien (Microsoft Corporation):
 *
 * BIFF (BInary File Format) is the file format in which Excel documents are
 * saved on disk.  A BIFF file is a complete description of an Excel document.
 * BIFF files consist of sequences of variable-length records. There are many
 * different types of BIFF records.  For example, one record type describes a
 * formula entered into a cell; one describes the size and location of a
 * window into a document; another describes a picture format.
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 * @category FileFormats
 * @package  Spreadsheet_Excel_Writer
 */

class BIFFwriter
{
    const BYTE_ORDER_LE = 0;
    const BYTE_ORDER_BE = 1;

    /**
     * @var integer
     */
    protected $version;

    /**
     * The byte order of this architecture. 0 => little endian, 1 => big endian
     * @var integer
     */
    protected $byteOrder;

    /**
     * The string containing the data of the BIFF stream
     * @var string
     */
    protected $data = '';

    /**
     * The size of the data in bytes. Should be the same as strlen($this->data)
     * @var integer
     */
    protected $datasize = 0;

    /**
     * The temporary dir for storing the OLE file
     * @var string
     */
    protected $tmpDir;

    /**
     * The temporary file for storing the OLE file
     * @var string
     */
    protected $tmpFile = '';

    /**
     * @var BiffInterface
     */
    protected $biff;

    /**
     * @param int $version
     *
     * @throws \Exception
     */
    public function __construct($version = Biff5::VERSION)
    {
        $this->tmpDir = sys_get_temp_dir();

        $this->setVersion($version);
        $this->setByteOrder();
    }

    /**
     * set BIFF version
     * @param $version
     *
     * @throws \Exception
     */
    protected function setVersion($version)
    {
        switch ($version) {
            case Biff5::VERSION:
                $this->biff = new Biff5;
                break;
            case Biff8::VERSION:
                $this->biff = new Biff8;
                break;
            default:
                throw new \Exception("Unsupported BIFF version");
        }

        $this->version = $version;
    }

    /**
     * Determine the byte order and store it as class data to avoid
     * recalculating it for each call to new().
     *
     */
    protected function setByteOrder()
    {
        // Check if "pack" gives the required IEEE 64bit float
        $teststr = pack("d", 1.2345);
        $number = pack("C8", 0x8D, 0x97, 0x6E, 0x12, 0x83, 0xC0, 0xF3, 0x3F);
        if ($number == $teststr) {
            $this->byteOrder = self::BYTE_ORDER_LE;
        } elseif ($number == strrev($teststr)) {
            $this->byteOrder = self::BYTE_ORDER_BE;
        } else {
            throw new \Exception(
                "Required floating point format is not supported on this platform."
            );
        }
    }

    /**
     * @param $data
     *
     * @return string
     */
    protected function addContinueIfNeeded($data)
    {
        if (strlen($data) > $this->biff->getLimit()) {
            $data = $this->addContinue($data);
        }

        return $data;
    }

    /**
     * @param string $data binary data to prepend
     */
    protected function prepend($data)
    {
        $data = $this->addContinueIfNeeded($data);

        $this->data = $data . $this->data;
        $this->datasize = strlen($this->data);
    }

    /**
     * @param string $data binary data to append
     */
    protected function append($data)
    {
        $data = $this->addContinueIfNeeded($data);

        $this->data = $this->data . $data;
        $this->datasize = strlen($this->data);
    }

    /**
     * Writes Excel BOF record to indicate the beginning of a stream or
     * sub-stream in the BIFF file.
     *
     * @param  integer $type Type of BIFF file to write: 0x0005 Workbook,
     *                       0x0010 Worksheet.
     * @throws \Exception
     */
    protected function storeBof($type)
    {
        $record = 0x0809; // Record identifier

        // According to the SDK $build and $year should be set to zero.
        // However, this throws a warning in Excel 5. So, use magic numbers.
        if ($this->isBiff5()) {
            $length = 0x0008;
            $unknown = '';
            $build = 0x096C;
            $year = 0x07C9;
        } else {
            $length = 0x0010;
            $unknown = pack("VV", 0x00000041, 0x00000006); //unknown last 8 bytes for BIFF8
            $build = 0x0DBB;
            $year = 0x07CC;
        }

        $header = pack("vv", $record, $length);
        $data = pack("vvvv", $this->version, $type, $build, $year);
        $this->prepend($header . $data . $unknown);
    }

    /**
     * Writes Excel EOF record to indicate the end of a BIFF stream.
     */
    protected function storeEof()
    {
        $record = 0x000A; // Record identifier
        $length = 0x0000; // Number of bytes to follow
        $header = pack("vv", $record, $length);
        $this->append($header);
    }

    /**
     * Excel limits the size of BIFF records. In Excel 5 the limit is 2084 bytes. In
     * Excel 97 the limit is 8228 bytes. Records that are longer than these limits
     * must be split up into CONTINUE blocks.
     *
     * This function takes a long BIFF record and inserts CONTINUE records as
     * necessary.
     *
     * @param  string $data The original binary data to be written
     * @return string        A very convenient string of continue blocks
     */
    protected function addContinue($data)
    {
        $limit = $this->biff->getLimit();
        $record = 0x003C; // Record identifier

        // The first 2080/8224 bytes remain intact. However, we have to change
        // the length field of the record.
        $tmp = substr($data, 0, 2) . pack("v", $limit - 4) . substr($data, 4, $limit - 4);

        $header = pack("vv", $record, $limit); // Headers for continue records

        // Retrieve chunks of 2080/8224 bytes +4 for the header.
        $dataLength = strlen($data);
        for ($i = $limit; $i < ($dataLength - $limit); $i += $limit) {
            $tmp .= $header;
            $tmp .= substr($data, $i, $limit);
        }

        // Retrieve the last chunk of data
        $header = pack("vv", $record, strlen($data) - $i);
        $tmp .= $header;
        $tmp .= substr($data, $i, strlen($data) - $i);

        return $tmp;
    }

    /**
     *
     */
    public function isBiff5()
    {
        return $this->version === Biff5::VERSION;
    }

    /**
     *
     */
    public function isBiff8()
    {
        return $this->version === Biff8::VERSION;
    }

    /**
     * @return int
     */
    public function getDataSize()
    {
        return $this->datasize;
    }
}
