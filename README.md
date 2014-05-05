PHP_XLSXWriter_plus
==============

This library is designed to be lightweight, and have relatively low memory usage. This is the fork of https://github.com/mk-j/PHP_XLSXWriter

It is designed to output an Excel spreadsheet in with (Office 2007+) xlsx format, with just basic features supported:
* assumes input is valid UTF-8
* multiple worksheets
* supports cell formats:
```
// 1 0
// 2 0.00
// 3 #,##0
// 4 #,##0.00
// 5 $#,##0_);($#,##0)
// 6 $#,##0_);[Red]($#,##0)
// 7 $#,##0.00_);($#,##0.00)
// 8 $#,##0.00_);[Red]($#,##0.00)
// 9 0%
// 10 0.00%
// 11 0.00E+00
// 12 # ?/?
// 13 # ??/??
// 14 m/d/yyyy
// 15 d-mmm-yy
// 16 d-mmm
// 17 mmm-yy
// 18 h:mm AM/PM
// 19 h:mm:ss AM/PM
// 20 h:mm
// 21 h:mm:ss
// 22 m/d/yyyy h:mm
// 37 #,##0_);(#,##0)
// 38 #,##0_);[Red](#,##0)
// 39 #,##0.00_);(#,##0.00)
// 40 #,##0.00_);[Red](#,##0.00)
// 45 mm:ss
// 46 [h]:mm:ss
// 47 mm:ss.0
// 48 ##0.0E+0
// 49 @
```
* supports styling cells


Simple example:
```php
$data = array(
    array('year','month','amount'),
    array('2003','1','220'),
    array('2003','2','153.5'),
);

$writer = new XLSXWriter();
$writer->writeSheet($data);
$writer->writeToFile('output.xlsx');
```

Multiple Sheets:
```php
$data1 = array(  
     array('5','3'),
     array('1','6'),
);
$data2 = array(  
     array('2','7','9'),
     array('4','8','0'),
);

$writer = new XLSXWriter();
$writer->setAuthor('Doc Author');
$writer->writeSheet($data1);
$writer->writeSheet($data2);
echo $writer->writeToString();
```

Cell Formatting:
```php
$header = array(
  'create_date'=>'string',
  'quantity'=>'string',
  'product_id'=>'string',
  'amount'=>'string',
  'description'=>'string',
);
$data = array(
    array('2013-01-01',1,27,'44.00','twig'),
    array('2013-01-05',1,'=C1','-44.00','refund'),
);

$writer = new XLSXWriter();
$writer->writeSheet($data,'Sheet1', $header);
$writer->writeToFile('example.xlsx');
```

Cell Styling
```php
    $writer->setFontName('MS Sans Serif'); //default document font name
    $writer->setFontSize(8); //default document font size
    $writer->setWrapText(true); //default document wrap cells
    $writer->setVerticalAlign('top'); //default document vertical align
    $writer->setHorizontalAlign('left'); //default document horizontal alilgn
    $writer->setStartRow(10); //set start data filling row
    $writer->setStartCol(0); //set start data filling column
    $writer->allDataFilledStyleFirst(false); //if true - 1st element of array will be used for styling all data filled cells

    //set styles
    $writer->setStyle(
      array (
        array( // in each style element you can use or 'cells', or 'rows' or 'columns'.
          'font' => array(
            'name' => 'Times New Roman',
            'size' => '11',
            'color' => '0000FF',
            'bold' => true,
            'italic' => true,
            'underline' => true),
          'border' => array(
            'style' => 'thin',
            'color' => 'A0A0A0'),
          'fill' => array(
            'color' => 'F0F0F0'),
          'cells' => array( //for 1 cell - array is not nessesary, use - 'cells' => 'C3'
            'E1',
            'E2'),
          'wrapText' => true,
          'verticalAlign' => 'top',
          'horizontalAlign' => 'left',
          'format' => 5
          ),
          array(
          'fill' => array(
            'color' => 'F09900'),
          'columns' => '2' //for only one and firs column dont use 'columns' => '0', use 'columns' => array('0')
          ),
          array(
          'fill' => array(
            'color' => 'F000F0'),
          'rows' => array(
            '0','1') //for only one and firs row dont use 'rows' => '0', use 'rows' => array('0')
          )
        )
      );
```