# xls-writer
Port of PEAR Spreadsheet Excel Writer.

This package is compliant with [PSR-0](http://www.php-fig.org/psr/0/), [PSR-1](http://www.php-fig.org/psr/1/), and [PSR-2](http://www.php-fig.org/psr/2/).
If you notice compliance oversights, please send a patch via pull request.

## Usage

### Write spreadsheet to file
```php
    require('vendor/autoload.php');

    // We give the path to our file here
    $workbook = new Xls\Writer('test.xls');

    $worksheet = $workbook->addWorksheet('My first worksheet');

    $worksheet->write(0, 0, 'Name');
    $worksheet->write(0, 1, 'Age');
    $worksheet->write(1, 0, 'John Smith');
    $worksheet->write(1, 1, 30);
    $worksheet->write(2, 0, 'Johann Schmidt');
    $worksheet->write(2, 1, 31);
    $worksheet->write(3, 0, 'Juan Herrera');
    $worksheet->write(3, 1, 32);

    // We still need to explicitly close the workbook
    $workbook->close();
```

## License
This library is released under [MIT](http://www.tldrlegal.com/license/mit-license) license.