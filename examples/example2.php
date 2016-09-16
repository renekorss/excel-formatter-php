<?php
/**
 * PHPExcelFormatter example 2
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

    // Output fields array. Key is column number starting from 0
    $formatterColumns = array(
        0 => 'username',
        2 => 'phone_no',
        4 => 'sex'
    );

    // Set our fields
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
