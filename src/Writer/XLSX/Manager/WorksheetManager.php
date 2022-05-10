<?php

namespace OpenSpout\Writer\XLSX\Manager;

use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\InvalidArgumentException;
use OpenSpout\Common\Exception\IOException;
use OpenSpout\Common\Helper\Escaper\XLSX as XLSXEscaper;
use OpenSpout\Common\Helper\StringHelper;
use OpenSpout\Common\Manager\OptionsManagerInterface;
use OpenSpout\Writer\Common\Entity\Options;
use OpenSpout\Writer\Common\Entity\Worksheet;
use OpenSpout\Writer\Common\Helper\CellHelper;
use OpenSpout\Writer\Common\Manager\ManagesCellSize;
use OpenSpout\Writer\Common\Manager\RowManager;
use OpenSpout\Writer\Common\Manager\Style\StyleMerger;
use OpenSpout\Writer\Common\Manager\WorksheetManagerInterface;
use OpenSpout\Writer\XLSX\Helper\DateHelper;
use OpenSpout\Writer\XLSX\Manager\Style\StyleManager;

/**
 * XLSX worksheet manager, providing the interfaces to work with XLSX worksheets.
 */
class WorksheetManager implements WorksheetManagerInterface
{
    use ManagesCellSize;

    /**
     * Maximum number of characters a cell can contain.
     *
     * @see https://support.office.com/en-us/article/Excel-specifications-and-limits-16c69c74-3d6a-4aaf-ba35-e6eb276e8eaa [Excel 2007]
     * @see https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3 [Excel 2010]
     * @see https://support.office.com/en-us/article/Excel-specifications-and-limits-ca36e2dc-1f09-4620-b726-67c00b05040f [Excel 2013/2016]
     */
    public const MAX_CHARACTERS_PER_CELL = 32767;

    public const SHEET_XML_FILE_HEADER = <<<'EOD'
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        EOD;

    /** @var bool Whether inline or shared strings should be used */
    protected $shouldUseInlineStrings;

    /** @var OptionsManagerInterface */
    private $optionsManager;

    /** @var RowManager Manages rows */
    private $rowManager;

    /** @var StyleManager Manages styles */
    private $styleManager;

    /** @var StyleMerger Helper to merge styles together */
    private $styleMerger;

    /** @var SharedStringsManager Helper to write shared strings */
    private $sharedStringsManager;

    /** @var XLSXEscaper Strings escaper */
    private $stringsEscaper;

    /** @var StringHelper String helper */
    private $stringHelper;

    /** @var array */
    private $columnLettersCache = [];

    /** @var array */
    private $registeredStylesCache = [];

    /** @var array */
    private $shouldApplyStyleOnEmptyCellCache = [];

    /** @var int */
    private $sheetOutlineLevelRow = 0;

    /** @var int */
    private $sheetOutlineLevelRowFileOffset;

    /** @var bool */
    private $showRowOutlineSummaryBelow;

    /**
     * WorksheetManager constructor.
     */
    public function __construct(
        OptionsManagerInterface $optionsManager,
        RowManager $rowManager,
        StyleManager $styleManager,
        StyleMerger $styleMerger,
        SharedStringsManager $sharedStringsManager,
        XLSXEscaper $stringsEscaper,
        StringHelper $stringHelper
    ) {
        $this->optionsManager = $optionsManager;
        $this->shouldUseInlineStrings = $optionsManager->getOption(Options::SHOULD_USE_INLINE_STRINGS);
        $this->setDefaultColumnWidth($optionsManager->getOption(Options::DEFAULT_COLUMN_WIDTH));
        $this->setDefaultRowHeight($optionsManager->getOption(Options::DEFAULT_ROW_HEIGHT));
        $this->columnWidths = $optionsManager->getOption(Options::COLUMN_WIDTHS) ?? [];
        $this->showRowOutlineSummaryBelow = $optionsManager->getOption(Options::SHOW_ROW_OUTLINE_SUMMARY_BELOW);
        $this->rowManager = $rowManager;
        $this->styleManager = $styleManager;
        $this->styleMerger = $styleMerger;
        $this->sharedStringsManager = $sharedStringsManager;
        $this->stringsEscaper = $stringsEscaper;
        $this->stringHelper = $stringHelper;
    }

    /**
     * @return SharedStringsManager
     */
    public function getSharedStringsManager()
    {
        return $this->sharedStringsManager;
    }

    /**
     * {@inheritdoc}
     */
    public function startSheet(Worksheet $worksheet)
    {
        $sheetFilePointer = fopen($worksheet->getFilePath(), 'w');
        $this->throwIfSheetFilePointerIsNotAvailable($sheetFilePointer);

        $worksheet->setFilePointer($sheetFilePointer);

        fwrite($sheetFilePointer, self::SHEET_XML_FILE_HEADER);
    }

    /**
     * {@inheritdoc}
     */
    public function addRow(Worksheet $worksheet, Row $row)
    {
        if (!$this->rowManager->isEmpty($row) || $row->getOutlineLevel() > 0) {
            $this->addNonEmptyRow($worksheet, $row);
        }

        $worksheet->setLastWrittenRowIndex($worksheet->getLastWrittenRowIndex() + 1);
    }

    /**
     * Write SheetPr.
     */
    private function writeSheetPr(Worksheet $worksheet)
    {
        $xml = '<sheetPr><outlinePr summaryBelow="'.($this->showRowOutlineSummaryBelow ? '1' : '0').'"/></sheetPr>';
        fwrite($worksheet->getFilePointer(), $xml);
    }

    /**
     * Write SheetViews.
     */
    private function writeSheetViews(Worksheet $worksheet)
    {
        $sheet = $worksheet->getExternalSheet();
        if ($sheet->hasSheetView()) {
            $xml = '<sheetViews>'.$sheet->getSheetView()->getXml().'</sheetViews>';
            fwrite($worksheet->getFilePointer(), $xml);
        }
    }

    /**
     * Write SheetFormatPr.
     */
    private function writeSheetFormatPr(Worksheet $worksheet)
    {
        $worksheetFilePointer = $worksheet->getFilePointer();

        $rowHeightXml = empty($this->defaultRowHeight)
            ? ' defaultRowHeight="0"'
            : " defaultRowHeight=\"{$this->defaultRowHeight}\"";
        $colWidthXml = empty($this->defaultColumnWidth)
            ? ''
            : " defaultColWidth=\"{$this->defaultColumnWidth}\"";

        $xml = "<sheetFormatPr{$colWidthXml}{$rowHeightXml} outlineLevelRow=\"0\"/>";

        $this->sheetOutlineLevelRowFileOffset = ftell($worksheetFilePointer) +
            strpos($xml, ' outlineLevelRow="') + 18;

        fwrite($worksheetFilePointer, $xml);
    }

    /**
     * Write Cols.
     */
    private function writeCols(Worksheet $worksheet)
    {
        if (empty($this->columnWidths)) {
            return;
        }

        $xml = '<cols>';
        foreach ($this->columnWidths as $entry) {
            $xml .= '<col min="'.$entry[0].'" max="'.$entry[1].'" width="'.$entry[2].'" customWidth="true"/>';
        }
        $xml .= '</cols>';

        fwrite($worksheet->getFilePointer(), $xml);
    }

    /**
     * Write MergeCells.
     */
    private function writeMergeCells(Worksheet $worksheet)
    {
        $mergeCells = $this->optionsManager->getOption(Options::MERGE_CELLS);
        if ($mergeCells) {
            $xml = '<mergeCells count="'.\count($mergeCells).'">';
            foreach ($mergeCells as $values) {
                $output = array_map(function ($value) {
                    return CellHelper::getColumnLettersFromColumnIndex($value[0]).$value[1];
                }, $values);
                $xml .= '<mergeCell ref="'.implode(':', $output).'"/>';
            }
            $xml .= '</mergeCells>';
            fwrite($worksheet->getFilePointer(), $xml);
        }
    }

    /**
     * Update the outlineLevel attribute in the sheetFormatPr element.
     */
    private function updateSheetOutlineLevelRow(Worksheet $worksheet, int $newOutlineLevel)
    {
        $this->sheetOutlineLevelRow = $newOutlineLevel;

        $worksheetFilePointer = $worksheet->getFilePointer();
        $oldFileOffset = ftell($worksheetFilePointer);
        fseek($worksheetFilePointer, $this->sheetOutlineLevelRowFileOffset);
        fwrite($worksheetFilePointer, $newOutlineLevel, 1);
        fseek($worksheetFilePointer, $oldFileOffset);
    }

    /**
     * {@inheritdoc}
     */
    public function close(Worksheet $worksheet)
    {
        $worksheetFilePointer = $worksheet->getFilePointer();

        if (!\is_resource($worksheetFilePointer)) {
            return;
        }
        $this->ensureSheetDataStated($worksheet);
        fwrite($worksheetFilePointer, '</sheetData>');

        $this->writeMergeCells($worksheet);

        fwrite($worksheetFilePointer, '</worksheet>');
        fclose($worksheetFilePointer);
    }

    /**
     * Writes the sheet data header.
     *
     * @param Worksheet $worksheet The worksheet to add the row to
     */
    private function ensureSheetDataStated(Worksheet $worksheet)
    {
        if (!$worksheet->getSheetDataStarted()) {
            $this->sheetOutlineLevelRow = 0;

            $this->writeSheetPr($worksheet);
            $this->writeSheetViews($worksheet);
            $this->writeSheetFormatPr($worksheet);
            $this->writeCols($worksheet);

            fwrite($worksheet->getFilePointer(), '<sheetData>');
            $worksheet->setSheetDataStarted(true);
        }
    }

    /**
     * Checks if the sheet has been sucessfully created. Throws an exception if not.
     *
     * @param bool|resource $sheetFilePointer Pointer to the sheet data file or FALSE if unable to open the file
     *
     * @throws IOException If the sheet data file cannot be opened for writing
     */
    private function throwIfSheetFilePointerIsNotAvailable($sheetFilePointer)
    {
        if (!$sheetFilePointer) {
            throw new IOException('Unable to open sheet for writing.');
        }
    }

    /**
     * Adds non empty row to the worksheet.
     *
     * @param Worksheet $worksheet The worksheet to add the row to
     * @param Row       $row       The row to be written
     *
     * @throws InvalidArgumentException If a cell value's type is not supported
     * @throws IOException              If the data cannot be written
     */
    private function addNonEmptyRow(Worksheet $worksheet, Row $row)
    {
        $this->ensureSheetDataStated($worksheet);
        $sheetFilePointer = $worksheet->getFilePointer();
        $rowStyle = $row->getStyle();
        $serializedRowStyle = $rowStyle->serialize();
        $rowIndexOneBased = $worksheet->getLastWrittenRowIndex() + 1;
        $numCells = $row->getNumCells();

        $hasCustomHeight = $this->defaultRowHeight > 0 ? '1' : '0';
        $rowOutlineLevel = $row->getOutlineLevel();

        $rowXML = "<row r=\"{$rowIndexOneBased}\" spans=\"1:{$numCells}\" customHeight=\"{$hasCustomHeight}\"";
        if ($rowOutlineLevel > 0) {
          $rowXML .= " outlineLevel=\"{$rowOutlineLevel}\"";
          if ($rowOutlineLevel > $this->sheetOutlineLevelRow) {
            $this->updateSheetOutlineLevelRow($worksheet, $rowOutlineLevel);
          }
        }
        $rowXML .= ">";

        foreach ($row->getCells() as $columnIndexZeroBased => $cell) {

            // Merging the cell style with its row style, applying and register it

            $cellStyle = $cell->getStyle();
            $serializedCellStyle = $cellStyle->serialize();

            if (!isset($this->registeredStylesCache[$serializedRowStyle][$serializedCellStyle])) {
                $mergedCellAndRowStyle = $this->styleMerger->merge($cellStyle, $rowStyle);

                $this->registeredStylesCache[$serializedRowStyle][$serializedCellStyle] =
                    $this->styleManager->registerStyle($mergedCellAndRowStyle);
            }

            $registeredStyle = $this->registeredStylesCache[$serializedRowStyle][$serializedCellStyle];

            // Generate the cell XML content

            $cellType = $cell->getType();
            $styleId = $registeredStyle->getId();

            if (!isset($this->columnLettersCache[$columnIndexZeroBased])) {
                $this->columnLettersCache[$columnIndexZeroBased] = CellHelper::getColumnLettersFromColumnIndex($columnIndexZeroBased);
            }
            $columnLetters = $this->columnLettersCache[$columnIndexZeroBased];

            $cellXML = '<c r="'.$columnLetters.$rowIndexOneBased.'"';
            $cellXML .= ' s="'.$styleId.'"';

            if ($cellType === Cell::TYPE_STRING) {
                $value = $cell->getValue();
                if (\strlen($value) > self::MAX_CHARACTERS_PER_CELL && $this->stringHelper->getStringLength($value) > self::MAX_CHARACTERS_PER_CELL) {
                    throw new InvalidArgumentException('Trying to add a value that exceeds the maximum number of characters allowed in a cell (32,767)');
                }

                if ($this->shouldUseInlineStrings) {
                    $cellXML .= ' t="inlineStr"><is><t>'.$this->stringsEscaper->escape($value).'</t></is></c>';
                } else {
                    $sharedStringId = $this->sharedStringsManager->writeString($value);
                    $cellXML .= ' t="s"><v>'.$sharedStringId.'</v></c>';
                }
            } elseif ($cellType === Cell::TYPE_BOOLEAN) {
                $cellXML .= ' t="b"><v>'.(int) ($cell->getValue()).'</v></c>';
            } elseif ($cellType === Cell::TYPE_NUMERIC) {
                $cellXML .= '><v>'.$this->stringHelper->formatNumericValue($cell->getValue()).'</v></c>';
            } elseif ($cellType === Cell::TYPE_FORMULA) {
                $cellXML .= '><f>'.substr($cell->getValue(), 1).'</f></c>';
            } elseif ($cellType === Cell::TYPE_DATE) {
                $value = $cell->getValue();
                if ($value instanceof \DateTimeInterface) {
                    $cellXML .= '><v>'.(string) DateHelper::toExcel($value).'</v></c>';
                } else {
                    throw new InvalidArgumentException('Trying to add a date value with an unsupported type: '.\gettype($value));
                }
            } elseif ($cellType === Cell::TYPE_ERROR && \is_string($cell->getValueEvenIfError())) {
                // only writes the error value if it's a string
                $cellXML .= ' t="e"><v>'.$cell->getValueEvenIfError().'</v></c>';
            } elseif ($cellType === Cell::TYPE_EMPTY) {
                if (!isset($this->shouldApplyStyleOnEmptyCellCache[$styleId])) {
                    $this->shouldApplyStyleOnEmptyCellCache[$styleId] = $this->styleManager->shouldApplyStyleOnEmptyCell($styleId);
                }

                if ($this->shouldApplyStyleOnEmptyCellCache[$styleId]) {
                    $cellXML .= '/>';
                } else {
                    // don't write empty cells that do no need styling
                    // NOTE: not appending to $cellXML is the right behavior!!
                    $cellXML = '';
                }
            } else {
                throw new InvalidArgumentException('Trying to add a value with an unsupported type: '.\gettype($cell->getValue()));
            }

            $rowXML .= $cellXML;
        }

        $rowXML .= '</row>';

        $wasWriteSuccessful = fwrite($sheetFilePointer, $rowXML);
        if (false === $wasWriteSuccessful) {
            throw new IOException("Unable to write data in {$worksheet->getFilePath()}");
        }
    }
}
