<?php

namespace Xls\Record;

class NameLong extends AbstractRecord
{
    const NAME = 'NAME';
    const ID = 0x0018;
    const LENGTH = 0x003d;

    /**
     * Store the NAME record in the long format that is used for storing the repeat
     * rows and columns when both are specified.
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
        $cce = 0x002e; // Length of text definition
        $ixals = $index + 1; // Sheet index
        $itab = $ixals; // Equal to ixals
        $cchCustMenu = 0x00; // Length of cust menu text
        $cchDescription = 0x00; // Length of description text
        $cchHelptopic = 0x00; // Length of help topic text
        $cchStatustext = 0x00; // Length of status bar text
        $rgch = $type; // Built-in name type

        $unknown01 = 0x29;
        $unknown02 = 0x002b;
        $unknown03 = 0x3b;
        $unknown04 = 0xffff - $index;
        $unknown05 = 0x0000;
        $unknown06 = 0x0000;
        $unknown07 = 0x1087;
        $unknown08 = 0x8008;

        $data = pack("v", $grbit);
        $data .= pack("C", $chKey);
        $data .= pack("C", $cch);
        $data .= pack("v", $cce);
        $data .= pack("v", $ixals);
        $data .= pack("v", $itab);
        $data .= pack("C", $cchCustMenu);
        $data .= pack("C", $cchDescription);
        $data .= pack("C", $cchHelptopic);
        $data .= pack("C", $cchStatustext);
        $data .= pack("C", $rgch);
        $data .= pack("C", $unknown01);
        $data .= pack("v", $unknown02);

        // Column definition
        $data .= pack("C", $unknown03);
        $data .= pack("v", $unknown04);
        $data .= pack("v", $unknown05);
        $data .= pack("v", $unknown06);
        $data .= pack("v", $unknown07);
        $data .= pack("v", $unknown08);
        $data .= pack("v", $index);
        $data .= pack("v", $index);
        $data .= pack("v", 0x0000);
        $data .= pack("v", 0x3fff);
        $data .= pack("C", $colmin);
        $data .= pack("C", $colmax);

        // Row definition
        $data .= pack("C", $unknown03);
        $data .= pack("v", $unknown04);
        $data .= pack("v", $unknown05);
        $data .= pack("v", $unknown06);
        $data .= pack("v", $unknown07);
        $data .= pack("v", $unknown08);
        $data .= pack("v", $index);
        $data .= pack("v", $index);
        $data .= pack("v", $rowmin);
        $data .= pack("v", $rowmax);
        $data .= pack("C", 0x00);
        $data .= pack("C", 0xff);

        // End of data
        $data .= pack("C", 0x10);

        return $this->getHeader() . $data;
    }
}
