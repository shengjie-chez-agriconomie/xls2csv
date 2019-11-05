<?php

namespace Component\FileHelper;

/**
 * Wrapper around fgetcsv for CSV sheets.
 */
class CsvReader implements RowReaderInterface
{
    /**
     * CSV file handle.
     *
     * @var resource file handle
     */
    private $fd;

    /**
     *  Field separator.
     *
     * @var string
     */
    private $separator;

    /*
     * Create reader instance for given CSV file
     * @param $filePath path to excel file
     * @param $separator field separator
     */

    /**
     * CsvReader constructor.
     *
     * @param $filePath
     * @param string $separator
     *
     * @throws \Exception
     */
    public function __construct($filePath, $separator = ',')
    {
        if (($fd = fopen($filePath, 'r')) === false) {
            throw new \Exception('Can\'t not open file : ' . $filePath);
        }

        $this->fd = $fd;
        $this->separator = $separator;
    }

    /**
     * Return next row as array or FALSE when end of file is reached
     * Behave the same with native fgetcsv function.
     *
     * @return array|false of string values
     */
    public function GetNextRow()
    {
        $row = fgetcsv($this->fd, 0, $this->separator);

        return $row;
    }

    /*
     * Close opened file and free its resources
     */

    public function Close()
    {
        if (isset($this->fd)) {
            fclose($this->fd);
        }
    }

    /**
     * {@inheritdoc}
     */
    public function setForceAllPages()
    {
        //only one page.
    }
}
