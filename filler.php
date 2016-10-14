<?php

/**
 * Auto word document filler
 * Tom Needham
 * tom@needham.im
 */

require 'vendor/autoload.php';
$phpWord = new \PhpOffice\PhpWord\PhpWord();

$options = getopts();

// Load the input template
if(!isset($options['t'])) {
    die('Need to supply input template (-t option)');
} else {
    $templatePath = $options['t'];
    if(!is_readable($templatePath)) {
        die('Cannot read the template file: '.$templatePath);
    }
}

// Load the data csv file
if(!isset($options['d'])) {
    die('Need to supply data .csv file (-d option)');
} else {
    $dataPath = $options['i'];
    if(!is_readable($dataPath)) {
        die('Cannot read the data file: '.$dataPath);
    }
}

// Check we can write to this directory
if(!is_writeable(__DIR__)) {
    die('Output path is not writeable');
}

// Load the template
$template = new \PhpOffice\PhpWord\TemplateProcessor($templatePath);

// Load the data file
$csvFile = fopen($dataPath, 'r');
$csv = fgetcsv($csvFile);

// Load the headers
$headers = $csv[0];

if(empty($headers)) {
    die('No variables found in the CSV header');
}

// Print out what variables we're replacing
echo "Found headers:\n"
foreach($headers as $header) {
    echo "-> $header\n";
}

// Check out OUTPUT column
if(!$outputColumn = array_search('OUTPUT', $headers)) {
    die('Must have an OUTPUT column in the CSV file to define the output filenames');
}

// How many rows have we got to do
$numFiles = count($csv)-1;
echo "$numFiles files to be created";

// Iterate through the rows with the data
for($i=1; $i<count($csv); $i++) {
    // Iterate through the variables to be replaced
    for($j=0; $j<count($headers); $j++) {
        $template->setValue($headers[$j], $csv[$i][$j]);
    }
    $fileName = $csv[$i][$outputColumn];
    if(!isset($options['p'])) {
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'PDF');
        $objWriter->save($fileName.'.pdf');
    } else {
        $template->saveAs($fileName.'.docx');
        echo "Saving file as: $fileName\n";
    }
}

echo 'Done!';

?>
