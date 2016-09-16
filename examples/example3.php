<?php
/**
 * PHPExcelFormatter example 3
 *
 * @author     Rene Korss <rene.korss@gmail.com>
 */

require __DIR__ . '/vendor/autoload.php';

use RKD\PHPExcelFormatter\PHPExcelFormatter;
use RKD\PHPExcelFormatter\Exception\PHPExcelFormatterException;

try
{
    // Load file
    $formatter = new PHPExcelFormatter('example2.xls', false);

    // Input columns array. Set column names for printing. Skip fourth column (third in array)
    $columns = array(
        'Username', 'E-mail', 'Phone', 4 => 'Sex'
    );

    // Output columns array
    $formatterColumns = array(
        'Username' => 'username',
        'Phone' => 'phone_no',
        'Sex' => 'sex'
    );

    // Set file columns, since first row is data, not field names
    $formatter->setColumns($columns);

    // Get file columns
    $fileColumns = $formatter->getColumns();

    // Print columns
    echo '<pre>'.print_r($fileColumns, true).'</pre>';

    // Set our columns
    $formatter->setFormatterColumns($formatterColumns);

    // Output as array
    $output = $formatter->output('a');

    // Print array
    echo '<pre>'.print_r($output, true).'</pre>';

    // Set MySQL table
    $formatter->setMySQLTableName('users');

    // Output as mysql query
    $output = $formatter->output('m');

    // Print mysql query
    echo '<pre>'.print_r($output, true).'</pre>';

}
catch(PHPExcelFormatterException $e)
{
    echo 'Error: '.$e->getMessage();
}

?>
