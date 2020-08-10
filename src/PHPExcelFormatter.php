<?php
/**
 * PHPExcelFormatter
 *
 * Copyright (c) 2020 Rene Korss
 *
 * @copyright  Copyright (c) 2020 Rene Korss
 * @license    http://opensource.org/licenses/MIT
 * @author     Rene Korss <rene.korss@gmail.com>
 */

namespace RKD\PHPExcelFormatter;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use RKD\PHPExcelFormatter\Exception\PHPExcelFormatterException;

/**
 * PHPExcelFormatter
 *
 * @copyright  Copyright (c) 2020 Rene Korss
 */

class PHPExcelFormatter
{
    /**
     * File for input
     */

    private $file;

    /**
     * Columns from excel file
     */

    private $columns = [];

    /**
     * Columns for output
     */

    private $formatterColumns = [];

    /**
     * MySQL table name
     */

    private $mysqlTable = '';

    /**
     * Highest row of worksheet
     */

    private $highestRow = 0;

    /**
     * Highest column of worksheet
     */

    private $highestColumn = 0;
    private $highestColumnIndex = 0;

    /**
     * Worksheet object
     */

    private $worksheetObj = null;

    /**
     * Field numbers by value
     */

    private $columnNumbers = [];

    /**
     * Row where to start to read file
     */

    private $startingRow = 1;

    /**
     * Constructor function
     *
     * @param   String  File path
     * @param   Boolean Do we read columns from first row
     *
     * @SuppressWarnings(PHPMD.BooleanArgumentFlag)
     */
    public function __construct($file, $readColumns = true)
    {
        // Check if we have PHPExcel
        if (!class_exists(Spreadsheet::class)) {
            throw new PHPExcelFormatterException(
                'Spreadsheet class not found. Please include it.'
            ); // @codeCoverageIgnore
        }

        // Set file
        $this->file = $file;

        // Create PHPExcel object
        $excelObj = new Spreadsheet();
        $inputFileType = IOFactory::identify($this->file);
        $readerObj = IOFactory::createReader($inputFileType);
        $readerObj->setReadDataOnly(true);

        // Load file to a PHPExcel Object
        $excelObj = $readerObj->load($this->file);

        // Set worksheet
        $this->worksheetObj = $excelObj->setActiveSheetIndex(0);

        $this->highestRow = $this->worksheetObj->getHighestRow();
        $this->highestColumn = $this->worksheetObj->getHighestColumn();
        $this->highestColumnIndex = Coordinate::columnIndexFromString($this->highestColumn);

        // If we need to read columns from first row
        if ($readColumns) {
            // If first row is columns, don't add it to formatted data
            $this->startingRow = 2;
            $row = 1;

            $columns = [];
            for ($col = 0; $col < $this->highestColumnIndex; ++$col) {
                $value = $this->worksheetObj->getCellByColumnAndRow($col + 1, $row, true)->getValue();
                $columns[$col] = $value;
                $this->columnNumbers[$value] = $col;
            }

            // Set columns
            $this->setColumns($columns);
        }
    }

    /**
     * Function to set excel columns
     *
     * @param   array   Columns
     */

    public function setColumns($columns = [])
    {
        $this->columns = (array)array_filter($columns);

        // Cell starts with 1, not 0
        $this->columnNumbers = array_map(function ($val) {
            return ++$val;
        }, array_flip($this->columns));
    }

    /**
     * Function to get columns
     *
     * @return  array   Array of columns
     */

    public function getColumns()
    {
        return $this->columns;
    }

    /**
     * Function to set formatter columns
     *
     * @param   array   Columns
     */

    public function setFormatterColumns($columns = [])
    {
        $this->formatterColumns = (array)$columns;
    }

    /**
     * Function to get formatter columns
     *
     * @return  array   Array of formatter columns
     */

    public function getFormatterColumns()
    {
        return $this->formatterColumns;
    }

    /**
     * Function to set MySQL table name
     *
     * @param   string  Table
     */

    public function setMySQLTableName($table)
    {
        $this->mysqlTable = $table;
    }

    /**
     * Function to get MySQL table name
     *
     * @return  string  MySQL table name
     */

    public function getMySQLTableName()
    {
        return $this->mysqlTable;
    }

    /**
     * Function to output formatted data
     *
     * @param   string  Format
     * @return  mixed   Array of results or MySQL query
     */

    public function output($format = '')
    {
        // Format data if not formated yet
        if (empty($this->formattedData)) {
            $this->format();
        }

        // Output depending on desired format
        switch ($format) {
            case 'a':
            case 'array':
                return $this->outputArray();
                break;
            case 'm':
            case 'mysql':
                return $this->outputMySQLQuery();
                break;
            default:
                return $this->outputMySQLQuery();
                break;
        }
    }

    /**
     * Function to format data
     *
     * @SuppressWarnings(PHPMD.ElseExpression)
     */
    public function format()
    {
        // Check if found column no
        if (empty($this->formatterColumns)) {
            throw new PHPExcelFormatterException('No formatter columns provided. Use setFormatterColumns() function.');
        }

        // Empty formatted data
        $this->formattedData = [];
        $formattedData = [];
        // Read all rows
        for ($row = $this->startingRow; $row <= $this->highestRow; ++$row) {
            foreach ($this->formatterColumns as $colIdentifier => $colName) {
                // If has column number
                if (is_int($colIdentifier) && $colIdentifier >= 0) {
                    // Cell numbring starts with 1
                    $colNo = $colIdentifier + 1;
                // If has column name
                } else {
                    // Check if we know columns
                    if (empty($this->columns)) {
                        // Columns are not
                        throw new PHPExcelFormatterException('Columns are not set.');
                    }

                    // Search column number for this column
                    $colNo = $this->columnNumbers[$colIdentifier] ?? false;

                    // Check if found column no
                    if (!(is_int($colNo) && $colNo >= 0)) {
                        throw new PHPExcelFormatterException('Field '.$colIdentifier.' not found.');
                    }
                }

                // Get value
                $value = $this->worksheetObj->getCellByColumnAndRow($colNo, $row)->getValue();

                // Set to formatted data with new name
                $formattedData[$row-1][$colName] = $value;
            }
        }

        $this->formattedData = $formattedData;
    }

    /**
     * Function to escape mysql column value
     *
     * Source: http://stackoverflow.com/a/1162502/1960712
     */

    protected function escape($value)
    {
        $search  = ["\\",  "\x00", "\n",  "\r",  "'",  '"', "\x1a"];
        $replace = ["\\\\","\\0","\\n", "\\r", "\'", '\"', "\\Z"];

        return str_replace($search, $replace, $value);
    }

    /**
     * OUTPUT FUNCTIONS
     */

    /**
     * Function to output data as array
     */

    protected function outputArray()
    {
        return $this->formattedData;
    }

    /**
     * Function to output data as MySQL query
     *
     * NB! This is not 100% secure against SQL injection. Should use outputArray() function with PDO or MySQLi.
     * Should support PDO and MySQLi. Remove function escape().
     */

    protected function outputMySQLQuery()
    {
        // Query
        $sql = '';

        // Sql rows
        $sqlRows = [];

        // Table name
        $tableName = $this->mysqlTable;

        // Get data
        $formattedData = $this->formattedData;

        // Check if we have MySQL table name
        if (strlen($tableName) == 0) {
            throw new PHPExcelFormatterException('MySQL table not set.');
        }

        // If we have data
        if (!empty($formattedData)) {
            // Create query
            $sql = "INSERT INTO `".$this->mysqlTable."` (`".implode('`, `', $this->formatterColumns)."`) VALUES ";

            foreach ($formattedData as $row) {
                // Start new row
                $sqlRow = '(';
                $sqlRowValues = null;

                foreach ($row as $columnValue) {
                    $sqlRowValues[] = "'".$this->escape($columnValue)."'";
                }

                // Add values
                $sqlRow .= implode(', ', $sqlRowValues);

                // End row
                $sqlRow .= ')';
                $sqlRows[] = $sqlRow;
            }

            $sql .= implode(', ', $sqlRows);
        }

        return $sql;
    }
}
