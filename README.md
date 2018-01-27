# Converter
文件转换类DocConverter，该类主要是吧office文件(word/excel/ppt)以及文本文件转为pdf文件，即把后缀为.doc、.docx、.xlx、.xlsx、.ppt、.pptx的office文件和.txt、.html、.js等文本文件转换为后缀为.pdf的文件。

###本类包含三个方法,第一个参数$srcfilename源文件的绝对路径（必选），第二个参数$destfilename目标文件的绝对路径（可选）：
* DoctPdf($srcfilename, $destfilename) //Word和文本文件转PDF；
* ExceltPdf($srcfilename, $destfilename) //Excel转PDF；
* PPTtPdf($srcfilename, $destfilename) //PPT转PDF。

##配置
### 1.php开启dcom扩展
* 打开php.ini，搜索php_com_dotnet和php_com_dotnet：
* extension=php_com_dotnet.dll   //把前面的分号去掉；
* com.allow_dcom = true  //改为true；
* 重启apache。

### 2.配置office组件服务
具体配置请看博客 -> [博客地址](http://blog.csdn.net/sangjinchao/article/details/78053545)