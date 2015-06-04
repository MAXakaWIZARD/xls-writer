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
    /**
     * This flag indicates write to temporary buffer mode
     * instead of $data
     * @var bool
     */
    protected $bufferedWrite = false;

    /**
     * Temporary buffer
     * @var string
     */
    protected $buffer = '';

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
     * @param $data
     *
     * @return string
     */
    protected function addContinueIfNeeded($data)
    {
        if (strlen($data) > Biff8::LIMIT) {
            $data = $this->addContinue($data);
        }

        return $data;
    }

    /**
     * @param string $data binary data to append
     */
    protected function append($data)
    {
        $data = $this->addContinueIfNeeded($data);
        $this->appendRaw($data);
    }

    /**
     * @param string $data binary data to append
     */
    protected function appendRaw($data)
    {
        if ($this->isBufferedWriteOn()) {
            $this->buffer .= $data;
        } else {
            $this->data .= $data;
            $this->datasize += strlen($data);
        }
    }

    /**
     * @param string $type
     * @param array $params
     */
    protected function appendRecord($type, array $params = array())
    {
        $this->append($this->getRecord($type, $params));
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
        return $this->getRecord('ContinueRecord', array($data));
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
        return new $className();
    }

    protected function isBufferedWriteOn()
    {
        return $this->bufferedWrite;
    }

    protected function startBufferedWrite()
    {
        $this->bufferedWrite = true;
        $this->buffer = '';
    }

    protected function endBufferedWrite()
    {
        $this->bufferedWrite = false;
    }

    protected function getBuffer()
    {
        return $this->buffer;
    }

    protected function getBufferSize()
    {
        return strlen($this->buffer);
    }

    /**
     * @return string
     */
    protected function getDataAndFlush()
    {
        $data = $this->data;
        $this->data = '';
        $this->datasize = 0;

        return $data;
    }
}
