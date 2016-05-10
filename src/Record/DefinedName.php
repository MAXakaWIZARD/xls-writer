<?php
namespace Xls\Record;

use Xls\StringUtils;

class DefinedName extends AbstractRecord
{
    const NAME = 'DEFINEDNAME';
    const ID = 0x18;

    const BUILTIN_PRINT_AREA = 0x06;
    const BUILTIN_PRINT_TITLES = 0x07;

    /**
     * @param $type
     * @param $sheetIndex
     * @param $formulaData
     *
     * @return string
     */
    public function getData($type, $sheetIndex, $formulaData)
    {
        $options = 0x20; // Option flags

        $name = pack("C", $type);
        $nameLen = StringUtils::countCharacters($name);
        $name = StringUtils::toBiff8UnicodeLongWoLenInfo($name);

        $formulaLen = strlen($formulaData);

        $data = pack("vC", $options, 0);
        $data .= pack("Cv", $nameLen, $formulaLen);
        $data .= pack("vv", 0, $sheetIndex);
        $data .= pack("CCCC", 0, 0, 0, 0);
        $data .= $name . $formulaData;

        return $this->getFullRecord($data);
    }
}
