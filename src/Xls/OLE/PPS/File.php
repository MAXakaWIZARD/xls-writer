<?php

namespace Xls\OLE\PPS;

use Xls\OLE;

/**
 * Class for creating File PPS's for OLE containers
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 * @category Structures
 * @package  OLE
 */
class File extends OLE\PPS
{
    /**
     * The constructor
     *
     * @param string $name The name of the file (in Unicode)
     * @see OLE::asc2Ucs()
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
     * @return mixed true on success.
     */
    public function init()
    {
        $this->tmpFilename = tempnam($this->tmpDir, "OLE_PPS_File");
        $this->ppsFile = @fopen($this->tmpFilename, "w+b");
        if ($this->ppsFile === false) {
            throw new \Exception("Can't create temporary file");
        }

        fseek($this->ppsFile, 0);

        return true;
    }

    /**
     * Append data to PPS
     *
     * @param string $data The data to append
     */
    public function append($data)
    {
        if (is_resource($this->ppsFile)) {
            fwrite($this->ppsFile, $data);
        } else {
            $this->data .= $data;
        }
    }

    /**
     * Returns a stream for reading this file using fread() etc.
     * @return  resource  a read-only stream
     */
    public function getStream()
    {
        $this->ole->getStream($this);
    }
}
