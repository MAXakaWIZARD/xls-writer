<?php

namespace Xls\Writer;

class Biff8 extends Biff5 implements BiffInterface
{
    /**
     * BIFF8
     *
     * Microsoft Excel 97 (XL8)
     * Microsoft Excel 2000 (XL9)
     * Microsoft Excel 2002 (XL10)
     * Microsoft Office Excel 2003 (XL11)
     * Microsoft Office Excel 2007 (XL12)
     */
    const VERSION = 0x0600;

    const MAX_ROWS = 65536;
    const MAX_COLS = 256;

    const MAX_STR_LENGTH = 255;
    const MAX_SHEET_NAME_LENGTH = 255;

    const LIMIT = 8228;
}
