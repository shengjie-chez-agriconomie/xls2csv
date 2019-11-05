<?php
require __DIR__.'/vendor/autoload.php';

require __DIR__.'/FileHelper/RowReaderInterface.php';
require __DIR__.'/FileHelper/RowWriterInterface.php';
require __DIR__.'/FileHelper/XlsWriter.php';
require __DIR__.'/FileHelper/CsvReader.php';

$file = $argv[1] ?? null;
$sep = $argv[2] ?? ',';
if (null === $file) {
	echo "usage csv2xls example.csv".PHP_EOL;
	exit;
}

if (!file_exists($file)) {
	echo "file {$file} not exists.".PHP_EOL;
	exit;
}


$r = new Component\FileHelper\CsvReader($file, $sep);
$w = new Component\FileHelper\XlsWriter(new \Liuggio\ExcelBundle\Factory(), $file.'.xls', 'xls');

while ($row = $r->GetNextRow()) {
	$w->WriteRow($row);
}

$r->Close();
$w->Close();

echo "file generated : ".$file.".xls".PHP_EOL;
