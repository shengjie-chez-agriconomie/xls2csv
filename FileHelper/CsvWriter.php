<?php

namespace Component\FileHelper;

/**
 * Wrapper around fputcsv for CSV sheets.
 */
class CsvWriter implements RowWriterInterface
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
     * Create writer instance for given CSV file
     * @param $filePath path to excel file
     * @param $separator field separator
     */

    /**
     * CsvWriter constructor.
     *
     * @param $filePath
     * @param string $separator
     *
     * @throws \Exception
     */
    public function __construct($filePath, $separator = ',')
    {
        if (($fd = fopen($filePath, 'w')) === false) {
            throw new \Exception('Can\'t not open file : ' . $filePath);
        }

        $this->fd = $fd;
        $this->separator = $separator;
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
        if (fputcsv($this->fd, $rowArray, $this->separator, '"') === false) {
            throw new \Exception('Can\'t write to file');
        }
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
}
