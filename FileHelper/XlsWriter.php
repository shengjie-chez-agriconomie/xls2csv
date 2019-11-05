<?php

namespace Component\FileHelper;

use Symfony\Component\DependencyInjection\ContainerInterface;

/**
 * Wrapper around for Excel acting like fputcsv.
 */
class XlsWriter implements RowWriterInterface
{
    /**
     *  phpexcel service.
     *
     * @var object class php excel service
     */
    private $phpExcelService;

    /**
     *  Excel-file instance.
     *
     * @var object class Excel object instance
     */
    private $phpExcelObject;

    /** Active sheet instance
     * @var object class active sheet
     */
    private $activeSheet;

    /**
     * Output file format (xls, xlsx, ...).
     *
     * @var string
     */
    private $format;

    /**
     *  File path.
     *
     * @var string
     */
    private $filePath;

    /**
     *  Current row index.
     *
     * @var int
     */
    private $currentRowIndex;

    /**
     * XlsWriter constructor.
     * Create writer instance for given Excel file path.
     *
     * @param ContainerInterface $container
     * @param string             $filePath  path to excel file
     * @param string             $format
     *
     * @throws \Exception
     */
    public function __construct($phpExcelService, $filePath, $format)
    {
        $phpExcelObject = $phpExcelService->createPHPExcelObject();
        if ($phpExcelObject == null) {
            throw new \Exception('Can\'t not create excel file in memory');
        }

        $this->phpExcelService = $phpExcelService;
        $this->phpExcelObject = $phpExcelObject;
        $this->activeSheet = $phpExcelObject->getActiveSheet();
        $this->currentRowIndex = 1;
        $this->format = $format;
        $this->filePath = $filePath;
    }

    /*
     * Append a new row at the end of the sheet
     *
     * @param $rowArray array
     */

    /**
     * @param $rowArray
     *
     * @throws \PHPExcel_Exception
     */
    public function WriteRow($rowArray)
    {
        $this->activeSheet->fromArray([$rowArray], null, 'A' . $this->currentRowIndex);

        ++$this->currentRowIndex;
    }

    /*
     * Close opened file and free its resources
     */
    public function Close()
    {
        /* @var \PHPExcel_Writer_Excel2007|\PHPExcel_Writer_Excel5 $writer */
        switch ($this->format) {
            case 'xls':
                $writer = $this->phpExcelService->createWriter($this->phpExcelObject, 'Excel5');
                break;
            case 'xlsx':
                $writer = $this->phpExcelService->createWriter($this->phpExcelObject, 'Excel2007');
                break;
            default:
                throw new \LogicException('XlsWriter::Close > unknown format ' . $this->format);
        }
        $writer->setPreCalculateFormulas(false);
        $writer->save($this->filePath);

        if (isset($this->phpExcelObject)) {
            $this->phpExcelObject->disconnectWorksheets();
            unset($this->phpExcelObject);
        }
    }

    /**
     * Add a new sheet in excel book and not move the pointer to it
     *
     * @param string $name
     */
    public function addSheet($name = 'N/A')
    {
        $newSheet = $this->phpExcelObject->createSheet();
        $newSheet->setTitle($name);
    }

    /**
     * Set active sheet to specific name
     *
     * @param string $sheetName
     */
    public function setActiveSheet($sheetName)
    {
        $this->activeSheet = $this->phpExcelObject->getSheetByName($sheetName);
        $this->currentRowIndex = 1; // If user want multiple manipulation with sheets, need refactor. Index back to 1 every change of active sheet
    }
}
