<?php

use PhpOffice\PhpWord\TemplateProcessor;

require_once __DIR__ . '/vendor/autoload.php';


$client = new \Google_Client();
$client->setApplicationName('Google Sheets API');
$client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
$client->setAccessType('offline');
$path = 'credentials.json';
$client->setAuthConfig($path);

$service = new \Google_Service_Sheets($client);
$spreadsheetId = '18U4J_Zzhanm-oangtg31VMcIF9ZuuGd5dYw_7vkNC2g';
$spreadsheet = $service->spreadsheets->get($spreadsheetId);


$rangeCoverPage = 'Cover Page'; // here we use the name of the Sheet to get all the rows
$response = $service->spreadsheets_values->get($spreadsheetId, $rangeCoverPage);
$valuesCoverPage = $response->getValues();

$rangeAbstract = 'Abstract'; // here we use the name of the Sheet to get all the rows
$response = $service->spreadsheets_values->get($spreadsheetId, $rangeAbstract);
$valuesAbstract = $response->getValues();

$rangeClassCodes = 'Classification Codes'; // here we use the name of the Sheet to get all the rows
$response = $service->spreadsheets_values->get($spreadsheetId, $rangeClassCodes);
$valuesClassCodes = $response->getValues();

$rangeClaims = 'Claims'; // here we use the name of the Sheet to get all the rows
$response = $service->spreadsheets_values->get($spreadsheetId, $rangeClaims);
$valuesClaims = $response->getValues();

$rangeApplication = 'Application Events'; // here we use the name of the Sheet to get all the rows
$response = $service->spreadsheets_values->get($spreadsheetId, $rangeApplication);
$valuesApplication = $response->getValues();

$count = 0;
$final = '';
foreach ($valuesClaims as $v){

    if ($count >2){
        foreach ($v as $a) {
            if (empty($a)){
                $final .= " \t ";
            }
            else{
                $final .= " ".$a;
            }
        }
        $final .= "\n";
    }

    $count++;
}

var_dump($valuesApplication);

$templateProcessor = new TemplateProcessor('template2.docx');
//var_dump($templateProcessor);

$templateProcessor->setValues([
    'title' => $valuesCoverPage[0][1],
    'patentNo' => $valuesCoverPage[1][1],
    'inventor' => $valuesCoverPage[2][1],
    'abstractTitle' => $valuesAbstract[3][1],
    'abstractText' => $valuesAbstract[4][1],
    'classCodeTitle' => $valuesClassCodes[0][1],
    'column1' => $valuesClassCodes[2][1],
    'column2' => $valuesClassCodes[2][2],
    'code1' => $valuesClassCodes[3][1],
    'code2' => $valuesClassCodes[4][1],
    'code3' => $valuesClassCodes[5][1],
    'code4' => $valuesClassCodes[6][1],
    'code5' => $valuesClassCodes[7][1],
    'desc1' => $valuesClassCodes[3][2],
    'desc2' => $valuesClassCodes[4][2],
    'desc3' => $valuesClassCodes[5][2],
    'desc4' => $valuesClassCodes[6][2],
    'desc5' => $valuesClassCodes[7][2],
    'claims' => $final,
    'appTitle' => $valuesApplication[1][0],
    'col3' => $valuesApplication[3][0],
    'col4' => $valuesApplication[3][1],
    'date1' => $valuesApplication[4][0],
    'status1' => $valuesApplication[4][1],
    'date2' => $valuesApplication[5][0],
    'status2' => $valuesApplication[5][1],
    'date3' => $valuesApplication[6][0],
    'status3' => $valuesApplication[6][1],
    'date4' => $valuesApplication[7][0],
    'status4' => $valuesApplication[7][1],
    'date5' => $valuesApplication[8][0],
    'status5' => $valuesApplication[8][1],
    'date6' => $valuesApplication[9][0],
    'status6' => $valuesApplication[9][1],
    'date7' => $valuesApplication[10][0],
    'status7' => $valuesApplication[10][1],
]);

$pathToSave = md5(time()).'.docx';
$templateProcessor->saveAs($pathToSave);

header('Content-Description: File Transfer');
header('Content-Disposition: attachment; filename=file.docx');
header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml/document');

readfile($pathToSave);