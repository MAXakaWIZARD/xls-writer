<?php

namespace Xls\Record;

class Password extends AbstractRecord
{
    const NAME = 'PASSWORD';
    const ID = 0x0013;
    const LENGTH = 0x02;

    /**
     * Generate the PASSWORD biff record
     * @param $password
     *
     * @return string
     */
    public function getData($password)
    {
        $data = pack("v", $password);

        return $this->getHeader() . $data;
    }
}
