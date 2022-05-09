<?php

namespace OpenSpout\Writer\Common\Entity;

/**
 * Writers' options holder.
 */
abstract class Options
{
    // CSV specific options
    public const FIELD_DELIMITER = 'fieldDelimiter';
    public const FIELD_ENCLOSURE = 'fieldEnclosure';
    public const SHOULD_ADD_BOM = 'shouldAddBOM';

    // Multisheets options
    public const TEMP_FOLDER = 'tempFolder';
    public const DEFAULT_ROW_STYLE = 'defaultRowStyle';
    public const SHOULD_CREATE_NEW_SHEETS_AUTOMATICALLY = 'shouldCreateNewSheetsAutomatically';
    public const SHOULD_APPLY_EXTRA_STYLES = 'shouldApplyExtraStyles';

    // XLSX specific options
    public const SHOULD_USE_INLINE_STRINGS = 'shouldUseInlineStrings';
    public const MERGE_CELLS = 'mergeCells';
    public const SHOW_ROW_OUTLINE_SUMMARY_BELOW = 'showRowOutlineSummaryBelow';

    // Cell size options
    public const DEFAULT_COLUMN_WIDTH = 'defaultColumnWidth';
    public const DEFAULT_ROW_HEIGHT = 'defaultRowHeight';
    public const COLUMN_WIDTHS = 'columnWidthDefinition';
}
