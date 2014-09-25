<?php
/**
 * PHPExcelFormatter
 *
 * Copyright (c) 2014 PHPExcelFormatter
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301
 * USA
 *
 * @category   PHPExcelFormatter
 * @package    PHPExcelFormatter
 * @copyright  Copyright (c) 2014 PHPExcelFormatter (https://github.com/renekorss/PHPExcelFormatter)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    1.0.0, 2014-09-24
 * @author     Rene Korss <rene.korss@gmail.com>
 */

// Require PHPExcel
require_once('libs/PHPExcel/PHPExcel.php');

/**
 * PHPExcelFormatterException
 */

class PHPExcelFormatterException extends Exception{ }

/**
 * PHPExcelFormatter
 *
 * @category   PHPExcelFormatter
 * @package    PHPExcelFormatter
 * @copyright  Copyright (c) 2014 PHPExcelFormatter (https://github.com/renekorss/PHPExcelFormatter)
 */

class PHPExcelFormatter
{
	/**
	 * File for input
	 */

	private $_file = '';

	/**
	 * Columns from excel file
	 */

	private $_columns = array();

	/**
	 * Columns for output
	 */

	private $_formatterColumns = array();

	/**
	 * MySQL table name
	 */

	private $_mysqlTable = '';

	/**
	 * Highest row of worksheet
	 */

	private $_highestRow = 0;

	/**
	 * Highest column of worksheet
	 */

	private $_highestColumn = 0;
	private $_highestColumnIndex = 0;

	/**
	 * Worksheet object
	 */

	private $_worksheetObj = null;

	/**
	 * Field numbers by value
	 */

	private $_columnNumbers = array();

	/**
	 * Row where to start to read file
	 */

	private $_startingRow = 1;

	/**
	 * Constructor function
	 *
	 * @param	String	File path
	 * @param	Boolean	Do we read columns from first row
	 */

	public function __construct($file, $readColumns = true)
	{
		// Set file
		$this->_file          = $file;

		// Initiate PHPExcel cache
		$cacheMethod               = PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
		$cacheSettings             = array('memoryCacheSize' => '32MB');
		PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

		// Create PHPExcel object
		$excelObj                  = new PHPExcel();
		$inputFileType             = PHPExcel_IOFactory::identify($this->_file);
		$readerObj                 = PHPExcel_IOFactory::createReader($inputFileType);
		$readerObj->setReadDataOnly(true);

		// Load file to a PHPExcel Object
		$excelObj                  = $readerObj->load($this->_file);

		// Set worksheet
		$this->_worksheetObj       = $excelObj->setActiveSheetIndex(0);

		$this->_highestRow         = $this->_worksheetObj->getHighestRow();
		$this->_highestColumn      = $this->_worksheetObj->getHighestColumn();
		$this->_highestColumnIndex = PHPExcel_Cell::columnIndexFromString($this->_highestColumn);

		// If we need to read columns from first row
		if($readColumns)
		{
			// If first row is columns, don't add it to formatted data
			$this->_startingRow = 2;
			$row = 1;

			for($col = 0; $col < $this->_highestColumnIndex; ++$col)
			{
				$value = $this->_worksheetObj->getCellByColumnAndRow($col, $row)->getValue();
				$columns[$col]                = $value;
				$this->_columnNumbers[$value] = $col;
			}

			// Set columns
			$this->setColumns($columns);
		}
	}

	/**
	 * Function to set excel columns
	 *
	 * @param	array	Columns
	 */

	public function setColumns($columns = array())
	{
		$this->_columns       = (array)$columns;
		$this->_columnNumbers = array_flip($this->_columns);
	}

	/**
	 * Function to get columns
	 *
	 * @return	array	Array of columns
	 */

	public function getColumns($columns = array())
	{
		return $this->_columns;
	}

	/**
	 * Function to set formatter columns
	 *
	 * @param	array	Columns
	 */

	public function setFormatterColumns($columns = array())
	{
		$this->_formatterColumns = (array)$columns;
	}

	/**
	 * Function to get formatter columns
	 *
	 * @return	array	Array of formatter columns
	 */

	public function getFormatterColumns()
	{
		return $this->_formatterColumns;
	}

	/**
	 * Function to set MySQL table name
	 *
	 * @param	string	Table
	 */

	public function setMySQLTableName($table = '')
	{
		$this->_mysqlTable = $table;
	}

	/**
	 * Function to get MySQL table name
	 *
	 * @return	string	MySQL table name
	 */

	public function getMySQLTableName()
	{
		return $this->_mysqlTable;
	}

	/**
	 * Function to output formatted data
	 *
	 * @param	string	Format
	 * @return	mixed	Array of results or MySQL query
	 */

	public function output($format = '')
	{
		// Format data if not formated yet
		if(empty($this->_formattedData))
		{
				$this->format();
		}

		// Output depending on desired format
		switch ($format)
		{
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
	 */

	public function format()
	{
		// Check if found column no
		if(empty($this->_formatterColumns))
			throw new PHPExcelFormatterException('No formatter columns provided. Use setFormatterColumns() function.');

		// Empty formatted data
		$this->_formattedData = array();

		// Read all rows
		for ($row = $this->_startingRow; $row <= $this->_highestRow; ++$row)
		{
			foreach ($this->_formatterColumns AS $colIdentifier => $colName)
			{
				// If has column number
				if(is_int($colIdentifier) AND $colIdentifier >= 0)
				{
					$colNo = $colIdentifier;
				}
				// If has column name
				else
				{
					// Check if we know columns
					if(empty($this->_columns))
					{
						// Columns are not
						throw new PHPExcelFormatterException('Columns are not set.');
					}

					// Search column number for this column
					$colNo = $this->_columnNumbers[$colIdentifier];

					// Check if found column no
					if(!(is_int($colNo) AND $colNo >= 0))
						throw new PHPExcelFormatterException('Field '.$colIdentifier.' not found.');
				}

				// Get value
				$value = $this->_worksheetObj->getCellByColumnAndRow($colNo, $row)->getValue();

				// Set to formatted data with new name
				$formattedData[$row-1][$colName] = $value;
			}
		}

		$this->_formattedData = $formattedData;
	}

	/**
	 * Function to escape mysql column value
	 *
	 * Source: http://stackoverflow.com/a/1162502/1960712
	 */

	protected function escape($value)
	{
		$search  = array("\\",  "\x00", "\n",  "\r",  "'",  '"', "\x1a");
		$replace = array("\\\\","\\0","\\n", "\\r", "\'", '\"', "\\Z");

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
		return $this->_formattedData;
	}

	/**
	 * Function to output data as MySQL query
	 *
	 * NB! This is not 100% secure against SQL injection. Should use outputArray() function with PDO or MySQLi.
	 * TODO: Should set DB connection and then escape. Support PDO and MySQLi. Remove function escape().
	 */

	protected function outputMySQLQuery()
	{
		// Query
		$sql           = '';

		// Sql rows
		$sqlRows       = array();

		// Table name
		$tableName     = $this->_mysqlTable;

		// Get data
		$formattedData = $this->_formattedData;

		// Check if we have MySQL table name
		if(strlen($tableName) == 0)
			throw new PHPExcelFormatterException('MySQL table not set.');

		// If we have data
		if(!empty($formattedData))
		{
			// Create query
			$sql = "INSERT INTO `".$this->_mysqlTable."` (`".implode('`, `', $this->_formatterColumns)."`) VALUES ";

			foreach ($formattedData as $row)
			{
				// Start new row
				$sqlRow = '(';
				$sqlRowValues = null;

				foreach ($row as $columnName => $columnValue)
				{
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
?>
