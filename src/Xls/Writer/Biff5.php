<?php

namespace Xls\Writer;

class Biff5
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
}
