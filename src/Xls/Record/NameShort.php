<?php
namespace Xls\Record;

class NameShort extends AbstractRecord
{
    const NAME = 'NAME';
    const ID = 0x0018;
    const LENGTH = 0x24;

    // Length of text definition
    const CCE = 0x0015;

    const UNKNOWN_08 = 0x8005;

    /**
     * Store the NAME record in the short format that is used for storing the print
     * area, repeat rows only and repeat columns only.
     *
     * @param integer $index  Sheet index
     * @param integer $type   Built-in name type
     * @param integer $rowmin Start row
     * @param integer $rowmax End row
     * @param integer $colmin Start colum
     * @param integer $colmax End column
     * @return string
     */
    public function getData($index, $type, $rowmin, $rowmax, $colmin, $colmax)
    {
        $grbit = 0x0020; // Option flags
        $chKey = 0x00; // Keyboard shortcut
        $cch = 0x01; // Length of text name
        $cce = static::CCE; // Length of text definition
        $ixals = $index + 1; // Sheet index
        $itab = $ixals; // Equal to ixals
        $cchCustMenu = 0x00; // Length of cust menu text
        $cchDescription = 0x00; // Length of description text
        $cchHelptopic = 0x00; // Length of help topic text
        $cchStatustext = 0x00; // Length of status bar text

        $data = pack("v", $grbit);
        $data .= pack("CCv", $chKey, $cch, $cce);
        $data .= pack("vv", $ixals, $itab);
        $data .= pack("C", $cchCustMenu);
        $data .= pack("C", $cchDescription);
        $data .= pack("C", $cchHelptopic);
        $data .= pack("C", $cchStatustext);
        $data .= pack("C", $type);

        $data .= $this->getExtraData($index, $rowmin, $rowmax, $colmin, $colmax);

        return $this->getHeader() . $data;
    }

    /**
     * @param $index
     * @param $rowmin
     * @param $rowmax
     * @param $colmin
     * @param $colmax
     *
     * @return string
     */
    protected function getExtraData($index, $rowmin, $rowmax, $colmin, $colmax)
    {
        $data = $this->getRowColDefCommonData($index);
        $data .= pack("vv", $rowmin, $rowmax);
        $data .= pack("CC", $colmin, $colmax);

        return $data;
    }

    /**
     * @param $index
     *
     * @return string
     */
    protected function getRowColDefCommonData($index)
    {
        $unknown03 = 0x3b;
        $unknown04 = 0xffff - $index;
        $unknown05 = 0x0000;
        $unknown06 = 0x0000;
        $unknown07 = 0x1087;
        $unknown08 = static::UNKNOWN_08;

        return pack(
            'Cvvvvvvv',
            $unknown03,
            $unknown04,
            $unknown05,
            $unknown06,
            $unknown07,
            $unknown08,
            $index,
            $index
        );
    }
}
