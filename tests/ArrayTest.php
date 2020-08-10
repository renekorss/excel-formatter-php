<?php
use PHPUnit\Framework\TestCase;
use RKD\PHPExcelFormatter\PHPExcelFormatter;

final class ArrayTest extends TestCase
{
    public function testCanRenameColumnsOnFirstRow(): void
    {
        // Load file
        $formatter = new PHPExcelFormatter(dirname(__DIR__).'/examples/example1.xls');

        // Output columns array
        $formatterColumns = [
            'username' => 'username',
            'phone'    => 'phone_no',
            'email'    => 'email_address'
        ];

        // Set our columns
        $formatter->setFormatterColumns($formatterColumns);

        // Output as array
        $output = $formatter->output('a');

        $firstRow = $output[1];

        // Renaming worked
        $this->assertArrayHasKey('phone_no', $firstRow);
        $this->assertArrayHasKey('email_address', $firstRow);

        // Got correct data
        $this->assertEquals('user', $firstRow['username']);
        $this->assertEquals(55555555, $firstRow['phone_no']);
        $this->assertEquals('user.name@gmail.com', $firstRow['email_address']);
    }

    public function testSetColumnNamesByIndex()
    {
        // Load file
        $formatter = new PHPExcelFormatter(dirname(__DIR__).'/examples/example2.xls', false);

        // Output fields array. Key is column number starting from 0
        $formatterColumns = [
            0 => 'username',
            2 => 'phone_no',
            4 => 'sex'
        ];

        // Set our fields
        $formatter->setFormatterColumns($formatterColumns);

        // Output as array
        $output = $formatter->output('a');

        $firstRow = $output[0];

        // Got correct data
        $this->assertEquals('user', $firstRow['username']);
        $this->assertEquals(554678876, $firstRow['phone_no']);
        $this->assertEquals('male', $firstRow['sex']);
    }

    public function testSetColumnNamesByStringAndIndex()
    {
        // Load file
        $formatter = new PHPExcelFormatter(dirname(__DIR__).'/examples/example2.xls', false);

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
        $output = $formatter->output('a');

        $firstRow = $output[0];

        // Got correct data
        $this->assertEquals('user', $firstRow['username']);
        $this->assertEquals(554678876, $firstRow['phone_no']);
        $this->assertEquals('male', $firstRow['sex']);
    }
}
