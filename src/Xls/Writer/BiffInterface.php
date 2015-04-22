<?php

namespace Xls\Writer;

interface BiffInterface
{
    /**
     * Returns the maximun length for a BIFF record
     * @return int
     */
    public function getLimit();

    /**
     * @return int
     */
    public function getCodepage();

    /**
     * @return int
     */
    public function getMaxSheetNameLength();
}
