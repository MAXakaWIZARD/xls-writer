<?php

namespace Xls\Writer;

class Biff5 implements BiffInterface
{
    /**
     * BIFF5
     *
     * Microsoft Excel version 5.0 (XL5)
     * Microsoft Excel 95 (XL7) (also called Microsoft Excel version 7)
     */
    const VERSION = 0x0500;

    const MAX_ROWS = 16384;
    const MAX_COLS = 256;

    const MAX_STR_LENGTH = 255;
    const MAX_SHEET_NAME_LENGTH = 31;

    const LIMIT = 2080;

    /**
     * The codepage indicates the text encoding used for strings
     */
    const CODEPAGE = 0x04E4;

    /**
     * @inheritdoc
     */
    public function getLimit()
    {
        return static::LIMIT;
    }

    /**
     * @inheritdoc
     */
    public function getCodepage()
    {
        return static::CODEPAGE;
    }

    /**
     * @inheritdoc
     */
    public function getMaxSheetNameLength()
    {
        return static::MAX_SHEET_NAME_LENGTH;
    }
}
