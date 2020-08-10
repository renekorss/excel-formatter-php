<?php
use PHPUnit\Framework\TestCase;
use RKD\PHPExcelFormatter\PHPExcelFormatter;

final class MySQLTest extends TestCase
{
    public function testCanRenameColumnsOnFirstRowForMysql(): void
    {
        // Load file
        $formatter = new PHPExcelFormatter(dirname(__DIR__).'/examples/example1.xls');
        $formatter->setMySQLTableName('users');

        // Output columns array
        $formatterColumns = [
            'username' => 'username',
            'phone'    => 'phone_no',
            'email'    => 'email_address'
        ];

        // Set our columns
        $formatter->setFormatterColumns($formatterColumns);

        // Output as array
        $output = $formatter->output('m');

        $this->assertEquals(
            "INSERT INTO `users` (`username`, `phone_no`, `email_address`) ".
            "VALUES ('user', '55555555', 'user.name@gmail.com'), ('test', '56789258', 'test@test.ee')",
            $output
        );
    }

    public function testSetColumnNamesByIndexForMysql()
    {
        // Load file
        $formatter = new PHPExcelFormatter(dirname(__DIR__).'/examples/example2.xls', false);
        $formatter->setMySQLTableName('users');

        // Output fields array. Key is column number starting from 0
        $formatterColumns = [
            0 => 'username',
            2 => 'phone_no',
            4 => 'sex'
        ];

        // Set our fields
        $formatter->setFormatterColumns($formatterColumns);

        // Output as array
        $output = $formatter->output('m');

        $this->assertEquals(
            "INSERT INTO `users` (`username`, `phone_no`, `sex`) ".
            "VALUES ('user', '554678876', 'male'), ('test', '428567867', 'female')",
            $output
        );
    }

    public function testSetColumnNamesByStringAndIndexForMysql()
    {
        // Load file
        $formatter = new PHPExcelFormatter(dirname(__DIR__).'/examples/example2.xls', false);
        $formatter->setMySQLTableName('users');

        // Input columns array. Set column names for printing. Skip fourth column (third in array)
        $columns = [
            'Username', 'E-mail', 'Phone', 4 => 'Sex'
        ];

        // Output columns array
        $formatterColumns = array(
            'Username' => 'username',
            'Phone' => 'phone_no',
            'Sex' => 'sex'
        );

        // Set file columns, since first row is data, not field names
        $formatter->setColumns($columns);

        // Set our fields
        $formatter->setFormatterColumns($formatterColumns);

        // Output as array
        $output = $formatter->output('m');

        $this->assertEquals(
            "INSERT INTO `users` (`username`, `phone_no`, `sex`) ".
            "VALUES ('user', '554678876', 'male'), ('test', '428567867', 'female')",
            $output
        );
    }
}
