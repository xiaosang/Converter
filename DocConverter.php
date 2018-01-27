<?php 
/**
 * 转换文件类
 * Class DocConverter
 * $srcfilename 源文件名（绝对路径）
 * $destfilename 目标文件名（绝对路径）
 */
class DocConverter {

    //文本文件转PDF,.doc、.docx、.txt等
    public function DoctPdf($srcfilename,$destfilename = '') {
    	if($destfilename == '') $destfilename = __DIR__ . '\DocConverter.pdf';
    	$srcfilename = str_replace('/', DIRECTORY_SEPARATOR , $srcfilename);
        $converttype = 0;
        try {
            if(!file_exists($srcfilename)){
                echo $srcfilename . ' is not exists';
                return;
            }
            $word = new \COM('word.application') or die("Can't start Word!");
            $word->Visible=0;
            $word->Documents->Open($srcfilename, false, false, false, '1', '1', true);
           
            $word->ActiveDocument->final = false;
            $word->ActiveDocument->Saved = true;
            $converttypetag;
            if ($converttype == 1) {
                $converttypetag = 2;        // wdExportCreateWordBookmarks
            } else {
                $converttypetag = 1;        // wdExportCreateHeadingBookmarks;
            }
            $word->ActiveDocument->ExportAsFixedFormat(
                $destfilename,
                17,                         // wdExportFormatPDF
                false,                      // open file after export
                0,                          // wdExportOptimizeForPrint
                3,                          // wdExportFromTo
                1,                          // begin page
                5000,                       // end page
                7,                          // wdExportDocumentWithMarkup
                true,                       // IncludeDocProps
                true,                       // KeepIRM
                $converttypetag             // WdExportCreateBookmarks
            );
            $word->ActiveDocument->Close();
            $word->Quit();
            echo 'topdf suceess:' . $destfilename;
        } catch (\Exception $e) {
            if (method_exists($word, 'Quit')){
                $word->Quit();
            }
            echo '[convert error]:' . $e->__toString();
            return;
        }
    }
    //Excel转PDF
    public function ExceltPdf($srcfilename,$destfilename = '') {
        if($destfilename == '') $destfilename = __DIR__ . '\EXcelConverter.pdf';
        $srcfilename = str_replace('/', DIRECTORY_SEPARATOR , $srcfilename);
        try {
            if(!file_exists($srcfilename)){
                echo $srcfilename . ' is not exists';
                return;
            }
            $excel = new \COM('excel.application') or die('Unable to instantiate excel');
            $workbook = $excel->Workbooks->Open($srcfilename, null, false, null, '1', '1', true);
            $workbook->ExportAsFixedFormat(0, $destfilename);
            $workbook->Close();
            $excel->Quit();
            echo 'topdf suceess:' . $destfilename;
        } catch (\Exception $e) {
            if (method_exists($excel, 'Quit')){
                $excel->Quit();
            }
            echo '[convert error]:' . $e->__toString();
            return;
        }
    }
    //PPT转PDF
    public function PPTtPdf($srcfilename,$destfilename = '') {
        if($destfilename == '') $destfilename = __DIR__ . '\PPTConverter.pdf';
        $srcfilename = str_replace('/', DIRECTORY_SEPARATOR , $srcfilename);
        try {
            if(!file_exists($srcfilename)){
                echo $srcfilename . ' is not exists';
                return;
            }
            $ppt = new \COM('powerpoint.application') or die('Unable to instantiate Powerpoint');
            $presentation = $ppt->Presentations->Open($srcfilename, false, false, false);
            $presentation->SaveAs($destfilename,32,1);
            $presentation->Close();
            $ppt->Quit();
            echo 'topdf suceess:' . $destfilename;
        } catch (\Exception $e) {
            if (method_exists($ppt, 'Quit')){
                $ppt->Quit();
            }
            echo '[convert error]:' . $e->__toString();
            return;
        }
    }

}