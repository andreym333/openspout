<?php

namespace OpenSpout\Common\Entity;

use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\InvalidArgumentException;

class Row
{
    /**
     * The cells in this row.
     *
     * @var Cell[]
     */
    protected $cells = [];

    /**
     * The row style.
     *
     * @var Style
     */
    protected $style;

    /**
     * Row height (default is 15).
     *
     * @var string
     */
    protected $height = '15';

    /**
     * Outline level.
     *
     * @var int
     */
    protected $outlineLevel = 0;

    /**
     * Row constructor.
     *
     * @param Cell[]     $cells
     * @param null|Style $style
     */
    public function __construct(array $cells, $style)
    {
        $this
            ->setCells($cells)
            ->setStyle($style)
        ;
    }

    /**
     * @return Cell[] $cells
     */
    public function getCells()
    {
        return $this->cells;
    }

    /**
     * @param Cell[] $cells
     *
     * @return Row
     */
    public function setCells(array $cells)
    {
        $this->cells = [];
        foreach ($cells as $cell) {
            $this->addCell($cell);
        }

        return $this;
    }

    /**
     * @param int $cellIndex
     *
     * @return Row
     */
    public function setCellAtIndex(Cell $cell, $cellIndex)
    {
        $this->cells[$cellIndex] = $cell;

        return $this;
    }

    /**
     * @param int $cellIndex
     *
     * @return null|Cell
     */
    public function getCellAtIndex($cellIndex)
    {
        return $this->cells[$cellIndex] ?? null;
    }

    /**
     * @return Row
     */
    public function addCell(Cell $cell)
    {
        $this->cells[] = $cell;

        return $this;
    }

    /**
     * @return int
     */
    public function getNumCells()
    {
        // When using "setCellAtIndex", it's possible to
        // have "$this->cells" contain holes.
        if (empty($this->cells)) {
            return 0;
        }

        return max(array_keys($this->cells)) + 1;
    }

    /**
     * @return Style
     */
    public function getStyle()
    {
        return $this->style;
    }

    /**
     * @param null|Style $style
     *
     * @return Row
     */
    public function setStyle($style)
    {
        $this->style = $style ?: new Style();

        return $this;
    }

    /**
     * @return array The row values, as array
     */
    public function toArray()
    {
        $values = [];
        foreach ($this->cells as $cellIndex => $cell) {
            $values[$cellIndex] = $cell->getValue();
        }
        return $values;
    }

    /**
     * Set row height.
     *
     * @param string $height
     *
     * @return Row
     */
    public function setHeight($height)
    {
        $this->height = $height;

        return $this;
    }

    /**
     * Returns row height.
     *
     * @return string
     */
    public function getHeight()
    {
        return $this->height;
    }

    /**
     * Get Outline Level.
     *
     * @return int
     */
    public function getOutlineLevel(): int
    {
        return $this->outlineLevel;
    }

    /**
     * Set Outline Level.
     *
     * @throws InvalidArgumentException If the value is not between 0 and 7
     *
     * @return Row
     */
    public function setOutlineLevel(int $level)
    {
        if ($level < 0 || $level > 7) {
            throw new InvalidArgumentException('Outline level must range between 0 and 7.');
        }

        $this->outlineLevel = $level;

        return $this;
    }
}
