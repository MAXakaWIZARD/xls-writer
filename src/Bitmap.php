<?php
namespace Xls;

class Bitmap
{
    const HEADER_SIZE = 0x36;

    /**
     * @var string
     */
    protected $filePath;

    /** Holds raw file data
     * @var mixed
     */
    protected $data;

    /**
     * @var mixed
     */
    protected $header;

    /**
     * @param string $filePath
     */
    public function __construct($filePath)
    {
        $this->filePath = $filePath;

        $this->read();
    }

    /**
     * @return string
     * @throws \Exception
     */
    protected function read()
    {
        $this->data = @file_get_contents($this->filePath, FILE_BINARY);
        if ($this->data === false) {
            throw new \Exception("Couldn't import $this->filePath");
        }

        $this->header = substr($this->data, 0, self::HEADER_SIZE);

        $this->validate();
    }

    /**
     * @throws \Exception
     */
    protected function validate()
    {
        if (strlen($this->data) <= self::HEADER_SIZE) {
            throw new \Exception("$this->filePath doesn't contain enough data");
        }

        if ($this->getIdent() != "BM") {
            throw new \Exception("$this->filePath doesn't appear to be a valid bitmap image");
        }

        if ($this->getColorDepth() != 24) {
            throw new \Exception("$this->filePath isn't a 24bit true color bitmap");
        }

        if ($this->getPlanesCount() != 1) {
            throw new \Exception("$this->filePath: only 1 plane supported in bitmap image");
        }

        if ($this->getCompression() != 0) {
            throw new \Exception("$this->filePath: compression not supported in bitmap image");
        }
    }

    /**
     * @return mixed
     */
    protected function getIdent()
    {
        $identity = unpack("A2ident", $this->header);

        return $identity['ident'];
    }

    /**
     * @return mixed
     */
    public function getWidth()
    {
        $result = unpack("Vwidth", substr($this->header, 18, 4));

        return $result['width'];
    }

    /**
     * @return mixed
     */
    public function getHeight()
    {
        $result = unpack("Vheight", substr($this->header, 22, 4));

        return $result['height'];
    }

    /**
     * @return mixed
     */
    public function getPlanesCount()
    {
        $result = unpack("v1", substr($this->header, 26, 2));

        return $result[1];
    }

    /**
     * @return mixed
     */
    public function getColorDepth()
    {
        $result = unpack("v1", substr($this->header, 28, 2));

        return $result[1];
    }

    /**
     * @return mixed
     */
    public function getCompression()
    {
        $result = unpack("V1", substr($this->header, 30, 4));

        return $result[1];
    }

    /**
     * @return string
     */
    public function getDataWithoutHeader()
    {
        return substr($this->data, self::HEADER_SIZE);
    }
}
