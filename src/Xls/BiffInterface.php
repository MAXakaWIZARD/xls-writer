<?php

namespace Xls;

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

    /**
     * @return string
     */
    public function getWorkbookName();

    /**
     * Returns length for a BOUNDSHEET record
     * @return int
     */
    public function getBoundsheetLength();
}
