<?php

namespace Xls\OLE;

/**
 * Class for creating File PPS's for OLE containers
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 * @category Structures
 * @package  OLE
 */
class PpsFile extends PPS
{
    /**
     * The constructor
     *
     * @param string $name The name of the file (in Unicode)
     */
    public function __construct($name)
    {
        parent::__construct(
            null,
            $name,
            OLE::PPS_TYPE_FILE
        );
    }

    /**
     * Initialization method. Has to be called right after OLE_PPS_File().
     * @throws \Exception
     * @return boolean true on success.
     */
    public function init()
    {
        $this->tmpFilename = tempnam($this->tmpDir, "OLE_PPS_File");
        $this->filePointer = @fopen($this->tmpFilename, "w+b");
        if ($this->filePointer === false) {
            throw new \Exception("Can't create temporary file");
        }

        fseek($this->filePointer, 0);

        return true;
    }

    /**
     * Append data to PPS
     *
     * @param string $data The data to append
     */
    public function append($data)
    {
        if (is_resource($this->filePointer)) {
            fwrite($this->filePointer, $data);
        }
    }
}