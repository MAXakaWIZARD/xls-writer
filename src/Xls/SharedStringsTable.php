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
}
