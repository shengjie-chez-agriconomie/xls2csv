<?php


namespace Component\FileHelper;


/**
 * Class HtmlTableWriter
 * @package Component\FileHelper
 */
class HtmlTableWriter implements RowWriterInterface
{
    /**
     * CSV file handle.
     *
     * @var resource file handle
     */
    private $fd;

    /*
     * Create writer instance for given CSV file
     * @param $filePath path to excel file
     * @param $separator field separator
     */

    /**
     * CsvWriter constructor.
     *
     * @param $filePath
     *
     * @throws \Exception
     */
    public function __construct($filePath)
    {
        if (($fd = fopen($filePath, 'w')) === false) {
            throw new \Exception('Can\'t not open file : ' . $filePath);
        }

        $this->fd = $fd;
        fwrite($this->fd,'<!DOCTYPE html><html><meta charset="UTF-8" /><body><table border="1">');
    }

    /*
     * Append a new row at the end of the sheet
     *
     * @param $rowArray array
     */

    /**
     * @param $rowArray
     *
     * @throws \Exception
     */
    public function WriteRow($rowArray)
    {
        fwrite($this->fd, '<tr>' . implode('', array_map(function ($item) {
                return '<td>' . htmlentities($item) . '</td>';
            }, $rowArray)) . '</tr>');
    }

    /*
     * Close opened file and free its resources
     */
    public function Close()
    {
        fwrite($this->fd,'</table></body></html>');
        if (isset($this->fd)) {
            fclose($this->fd);
        }
    }
}