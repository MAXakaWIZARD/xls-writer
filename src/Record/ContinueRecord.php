<?php

namespace Xls\Record;

use Xls\Biff8;

class ContinueRecord extends AbstractRecord
{
    const NAME = 'CONTINUE';
    const ID = 0x003C;

    /**
     * Excel limits the size of BIFF records. In Excel 97 the limit is 8228 bytes.
     * Records that are longer than these limits
     * must be split up into CONTINUE blocks.
     *
     * This function takes a long BIFF record and inserts CONTINUE records as
     * necessary.
     *
     * @param string $data The original binary data to be written
     *
     * @return string Сonvenient string of continue blocks
     */
    public function getData($data)
    {
        //reserve 4 bytes for header
        $limit = Biff8::LIMIT - 4;

        // The first bytes below limit remain intact. However, we have to change
        // the length field of the record.
        $recordId = substr($data, 0, 2);
        $newRecordSize = $limit - 4;
        $recordData = substr($data, 4, $newRecordSize);
        $result = $recordId . pack("v", $newRecordSize) . $recordData;

        $data = substr($data, $newRecordSize + 4);
        $result .= $this->getDataRaw($data);

        return $result;
    }

    public function getDataRaw($data)
    {
        //reserve 4 bytes for header
        $limit = Biff8::LIMIT - 4;

        $dataLength = strlen($data);

        $result = '';

        // Retrieve chunks of 8224 bytes +4 for the header
        for ($i = 0; $i < $dataLength - $limit; $i += $limit) {
            $chunk = substr($data, $i, $limit);
            $result .= $this->getFullRecord($chunk);
        }

        // Retrieve the last chunk of data
        $lastChunkLength = $dataLength - $i;
        if ($lastChunkLength > 0) {
            $lastChunk = substr($data, $i, $lastChunkLength);
            $result .= $this->getFullRecord($lastChunk);
        }

        return $result;
    }
}
