<?php

namespace Xls;

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

    const LIMIT = 2084;
    const BOUNDSHEET_LENGTH = 11;

    const WORKBOOK_NAME = 'Book';

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

    /**
     * @inheritdoc
     */
    public function getWorkbookName()
    {
        return static::WORKBOOK_NAME;
    }

    /**
     * @inheritdoc
     */
    public function getBoundsheetLength()
    {
        return static::BOUNDSHEET_LENGTH;
    }
}
