<?php

namespace Xls\OLE;

class PpsFile extends PPS
{
    /**
     * The constructor
     *
     * @param string $name The name of the file (in Unicode)
     */
    public function __construct($name)
    {
        parent::__construct(
            null,
            OLE::asc2Ucs($name),
            self::PPS_TYPE_FILE
        );
    }
}
