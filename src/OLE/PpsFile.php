<?php

namespace Xls\OLE;

/**
 * Class for creating File PPS's for OLE containers
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 * @category Structures
 * @package  OLE
 */
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
            OLE::PPS_TYPE_FILE
        );
    }
}
