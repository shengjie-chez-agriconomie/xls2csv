<?php
require __DIR__.'/vendor/autoload.php';

require __DIR__.'/FileHelper/RowReaderInterface.php';
require __DIR__.'/FileHelper/RowWriterInterface.php';
require __DIR__.'/FileHelper/XlsReader.php';
require __DIR__.'/FileHelper/CsvWriter.php';

$file = $argv[1] ?? null;
if (null === $file) {
	echo "usage xls2csv example.xls".PHP_EOL;
	exit;
}

if (!file_exists($file)) {
	echo "file {$file} not exists.".PHP_EOL;
	exit;
}


$r = new Component\FileHelper\XlsReader(new \Liuggio\ExcelBundle\Factory(), $file);
$w = new Component\FileHelper\CsvWriter($file.'.csv', ',');


while ($row = $r->GetNextRow()) {
	$w->WriteRow($row);
}

$r->Close();
$w->Close();

echo "file generated : ".$file.".csv".PHP_EOL;