<?php 
require_once('./DocConverter.php');
$doc = new DocConverter();
$doc->DoctPdf(__DIR__ . '/test.docx');
$doc->ExceltPdf(__DIR__ . '/test.xlsx');
$doc->PPTtPdf(__DIR__ . '/test.ppt');

 ?>