<?php

namespace Component\FileHelper;

/**
 * Interface for acting like fputcsv.
 */
interface RowWriterInterface
{
    /*
     * Append a new row at the end of the sheet
     *
     * @param $rowArray array
     */

    /**
     * @param $rowArray
     */
    public function WriteRow($rowArray);

    /*
     * Close opened file and free its resources
     */
    public function Close();
}
