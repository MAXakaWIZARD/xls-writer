<?php

namespace Xls\OLE;

use Xls\OLE\OLE;

/**
 * Stream wrapper for reading data stored in an OLE file. Implements methods
 * for PHP's stream_wrapper_register(). For creating streams using this
 * wrapper, use OLE\PPS\File::getStream().
 *
 * @category   Structures
 * @package    OLE
 * @author     Christian Schmidt <schmidt@php.net>
 * @license    http://www.php.net/license/3_0.txt  PHP License 3.0
 * @version    Release: @package_version@
 * @link       http://pear.php.net/package/OLE
 * @since      Class available since Release 0.6.0
 */
class ChainedBlockStream
{
    /**
     * The OLE container of the file that is being read.
     * @var OLE
     */
    public $ole;

    /**
     * Parameters specified by fopen().
     * @var array
     */
    public $params;

    /**
     * The binary data of the file.
     * @var  string
     */
    public $data;

    /**
     * The file pointer.
     * @var  int  byte offset
     */
    public $pos;

    /**
     * Implements support for fopen().
     * For creating streams using this wrapper, use OLE_PPS_File::getStream().
     * @param  string  $path resource name including scheme, e.g.
     *                 ole-chainedblockstream://oleInstanceId=1
     * @param  string  $mode only "r" is supported
     * @param  int     $options mask of STREAM_REPORT_ERRORS and STREAM_USE_PATH
     * @param  string  $openedPath absolute path of the opened stream (out parameter)
     * @throws \Exception
     * @return bool    true on success
     */
    public function stream_open($path, $mode, $options, &$openedPath)
    {
        if ($mode != 'r') {
            if ($options & STREAM_REPORT_ERRORS) {
                throw new \Exception('Only reading is supported', E_USER_WARNING);
            }

            return false;
        }

        // 25 is length of "ole-chainedblockstream://"
        parse_str(substr($path, 25), $this->params);
        if (!isset($this->params['oleInstanceId'],
                $this->params['blockId'],
                OLE::$instances[$this->params['oleInstanceId']]
            )
        ) {
            if ($options & STREAM_REPORT_ERRORS) {
                throw new \Exception('OLE stream not found', E_USER_WARNING);
            }

            return false;
        }
        $this->ole = OLE::$instances[$this->params['oleInstanceId']];

        $blockId = $this->params['blockId'];
        $this->data = '';

        $isSmallBlock = $this->isSmallBlock($blockId);
        $rootPos = ($isSmallBlock) ? $this->ole->getBlockOffset($this->ole->root->StartBlock) : 512;
        while ($blockId != -2) {
            $pos = $this->ole->getBlockOffset($blockId, $rootPos);
            fseek($this->ole->fileHandle, $pos);
            $this->data .= fread($this->ole->fileHandle, $this->ole->bigBlockSize);
            if ($isSmallBlock) {
                $blockId = $this->ole->sbat[$blockId];
            } else {
                $blockId = $this->ole->bbat[$blockId];
            }
        }

        if (isset($this->params['size'])) {
            $this->data = substr($this->data, 0, $this->params['size']);
        }

        if ($options & STREAM_USE_PATH) {
            $openedPath = $path;
        }

        return true;
    }

    /**
     * @param $blockId
     *
     * @return bool
     */
    protected function isSmallBlock($blockId)
    {
        return isset($this->params['size'])
            && $this->params['size'] < $this->ole->bigBlockThreshold
            && $blockId != $this->ole->root->StartBlock;
    }

    /**
     * Implements support for fclose().
     * @return  string
     */
    public function stream_close()
    {
        $this->ole = null;
        OLE::$instances = array();
    }

    /**
     * Implements support for fread(), fgets() etc.
     * @param   int  $count maximum number of bytes to read
     * @return  string
     */
    public function stream_read($count)
    {
        if ($this->stream_eof()) {
            return false;
        }
        $s = substr($this->data, $this->pos, $count);
        $this->pos += $count;
        return $s;
    }

    /**
     * Implements support for feof().
     * @return  bool  TRUE if the file pointer is at EOF; otherwise FALSE
     */
    public function stream_eof()
    {
        return $this->pos >= strlen($this->data);
    }

    /**
     * Returns the position of the file pointer, i.e. its offset into the file
     * stream. Implements support for ftell().
     * @return  int
     */
    public function stream_tell()
    {
        return $this->pos;
    }

    /**
     * Implements support for fseek().
     * @param   int $offset byte offset
     * @param   int $whence SEEK_SET, SEEK_CUR or SEEK_END
     * @return  bool
     */
    public function stream_seek($offset, $whence)
    {
        if ($whence == SEEK_SET && $offset >= 0) {
            $this->pos = $offset;
        } elseif ($whence == SEEK_CUR && -$offset <= $this->pos) {
            $this->pos += $offset;
        } elseif ($whence == SEEK_END && -$offset <= sizeof($this->data)) {
            $this->pos = strlen($this->data) + $offset;
        } else {
            return false;
        }

        return true;
    }

    /**
     * Implements support for fstat(). Currently the only supported field is
     * "size".
     * @return  array
     */
    public function stream_stat()
    {
        return array(
            'size' => strlen($this->data),
        );
    }

    // Methods used by stream_wrapper_register() that are not implemented:
    // bool stream_flush ( void )
    // int stream_write ( string data )
    // bool rename ( string path_from, string path_to )
    // bool mkdir ( string path, int mode, int options )
    // bool rmdir ( string path, int options )
    // bool dir_opendir ( string path, int options )
    // array url_stat ( string path, int flags )
    // string dir_readdir ( void )
    // bool dir_rewinddir ( void )
    // bool dir_closedir ( void )
}
