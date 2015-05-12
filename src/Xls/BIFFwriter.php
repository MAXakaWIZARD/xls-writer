<?php

namespace Xls;

use Xls\Record\AbstractRecord;

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

    const BOF_TYPE_WORKBOOK = 0x0005;
    const BOF_TYPE_WORKSHEET = 0x0010;

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
     * But this is not true for Worksheet, cause it writes directly to file
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
     * @return int
     */
    public function getVersion()
    {
        return $this->version;
    }

    /**
     * Determine the byte order and store it
     *
     */
    protected function setByteOrder()
    {
        // Check if "pack" gives the required IEEE 64bit float
        $teststr = pack("d", 1.2345);
        $number = pack("C8", 0x8D, 0x97, 0x6E, 0x12, 0x83, 0xC0, 0xF3, 0x3F);
        if ($number == $teststr) {
            $this->byteOrder = self::BYTE_ORDER_LE;
        } else {
            $this->byteOrder = self::BYTE_ORDER_BE;
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
        $this->datasize += strlen($data);
    }

    /**
     * @param string $data binary data to append
     */
    protected function append($data)
    {
        $data = $this->addContinueIfNeeded($data);

        $this->data .= $data;
        $this->datasize += strlen($data);
    }

    /**
     * @param string $type
     * @param array $params
     *
     * @return mixed
     */
    protected function appendRecord($type, array $params = array())
    {
        $this->append($this->getRecord($type, $params));
    }

    /**
     * @param string $type
     * @param array $params
     *
     * @return mixed
     */
    protected function prependRecord($type, array $params = array())
    {
        $this->prepend($this->getRecord($type, $params));
    }

    /**
     * This function takes a long BIFF record and inserts CONTINUE records as
     * necessary.
     *
     * @param  string $data The original binary data to be written
     * @return string Сonvenient string of continue blocks
     */
    protected function addContinue($data)
    {
        return $this->getRecord('ContinueRecord', array($data, $this->biff->getLimit()));
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

    /**
     * @param string $type
     * @param array $params
     *
     * @return mixed
     */
    protected function getRecord($type, array $params = array())
    {
        $record = $this->createRecord($type);

        return call_user_func_array(array($record, 'getData'), $params);
    }

    /**
     * @param $type
     *
     * @return AbstractRecord
     */
    protected function createRecord($type)
    {
        $className = "\\Xls\\Record\\$type";
        return new $className($this->version, $this->byteOrder);
    }
}
