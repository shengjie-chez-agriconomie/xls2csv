<?php

namespace Component\FileHelper;

use Symfony\Component\DependencyInjection\ContainerInterface;

/**
 * Wrapper around for Excel acting like fgetcsv.
 */
class XlsReader implements RowReaderInterface
{
    /**
     *  Excel-file instance.
     *
     * @var object class Excel object instance
     */
    private $phpExcelObject;

    /**
     *  Row iterator.
     *
     * @var object class row iterator
     */
    private $rowIterator;

    /**
     *  Current row index.
     *
     * @var int
     */
    private $currentRowIndex;

    /**
     *  Number of rows in spreadsheet.
     *
     * @var int
     */
    private $numberOrRows;

    /**
     * do we do only the current sheet ?
     *
     * @var bool
     */
    private $allPages = false;

    /**
     * XlsReader constructor.
     * Create reader instance for given Excel file path.
     *
     * @param ContainerInterface $container
     * @param string             $filePath  path to excel file
     *
     * @throws \Exception
     */
    public function __construct($factory, $filePath)
    {
        $phpExcelObject = $factory->createPHPExcelObject($filePath);
        if ($phpExcelObject == null) {
            throw new \Exception('Can\'t not open excel file : ' . $filePath);
        }

        $activeSheet = $phpExcelObject->getActiveSheet();
        $this->phpExcelObject = $phpExcelObject;
        $this->rowIterator = $activeSheet->getRowIterator(1);
        $this->currentRowIndex = 1;
        $this->numberOrRows = $activeSheet->getHighestRow();
    }

    /*
     * Return next row as array or FALSE when end of file is reached
     * Behave the same with native fgetcsv function
     *
     * @return array of string values
     */

    /**
     * @return array|bool
     *
     * @throws \PHPExcel_Calculation_Exception
     */
    public function GetNextRow()
    {
        if ($this->currentRowIndex > $this->numberOrRows) {
            if (!$this->allPages) {
                return false;
            }
            if ($this->phpExcelObject->getActiveSheetIndex() + 1 >= count($this->phpExcelObject->getAllSheets())) {
                return false;
            }

            $activeSheet = $this->phpExcelObject->getSheet($this->phpExcelObject->getActiveSheetIndex() + 1);

            $this->rowIterator = $activeSheet->getRowIterator(1);
            $this->currentRowIndex = 1;
            $this->numberOrRows = $activeSheet->getHighestRow();
        }

        $row = $this->rowIterator->current();

        /** @var \PHPExcel_Cell[]|\PHPExcel_Worksheet_CellIterator $cellIterator */
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false);
        $rowArray = [];
        foreach ($cellIterator as $cell) {
            $rowArray[] = strval($cell->getFormattedValue());
        }

        $this->rowIterator->next();
        ++$this->currentRowIndex;

        return $rowArray;
    }

    /*
     * Close opened file and free its resources
     */

    public function Close()
    {
        if (isset($this->rowIterator)) {
            unset($this->rowIterator);
        }
        if (isset($this->phpExcelObject)) {
            $this->phpExcelObject->disconnectWorksheets();
            unset($this->phpExcelObject);
        }
    }

    /**
     * {@inheritdoc}
     */
    public function setForceAllPages()
    {
        $this->allPages = true;

        $activeSheet = $this->phpExcelObject->getSheet(0);

        $this->rowIterator = $activeSheet->getRowIterator(1);
        $this->currentRowIndex = 1;
        $this->numberOrRows = $activeSheet->getHighestRow();
    }
}
