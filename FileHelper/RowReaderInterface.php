<?php

namespace Component\FileHelper;

/**
 * Interface for acting like fgetcsv.
 */
interface RowReaderInterface
{
    /*
     * Return next row as array or FALSE when end of file is reached
     * Behave the same with native fgetcsv function
     *
     * @return array of string values
     */

    public function GetNextRow();

    /*
     * Close opened file and free its resources
     */

    public function Close();

    /**
     * by default we only read the current pages,
     * setting this, will reset the cursor to the first and when current page is done,
     * we go to the next and the next until the final page.
     */
    public function setForceAllPages();
}
