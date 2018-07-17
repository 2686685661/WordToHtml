<?php

$rt = new Word2Json();

$fileName = __DIR__.DIRECTORY_SEPARATOR.'b'.DIRECTORY_SEPARATOR.'test.docx';
$res = $rt->readDocument($fileName);

// var_dump($res); 



class Word2Json
{
	private $rels_xml;
	private $doc_xml;
	
	/**
	 * READS The Document and Relationships into separated XML files
	 * 
	 * @param String $filename The filename
	 * @return void
	 */
	private function readZipPart($filename)
	{
		$zip = new ZipArchive();
		$_xml = 'word/document.xml';
		$_xml_rels = 'word/_rels/document.xml.rels';
		if (true === $zip->open($filename)) {
			if (($index = $zip->locateName($_xml)) !== false) {
				$xml = $zip->getFromIndex($index);
			}
			$zip->close();
		} else die('non zip file');
		
		if (true === $zip->open($filename)) {
			if (($index = $zip->locateName($_xml_rels)) !== false) {
				$xml_rels = $zip->getFromIndex($index);					
			}
			$zip->close();
		} else die('non zip file');

		// $a = $this->preprocessingWord($filename);
		// echo $a;die;
		
		$this->doc_xml = new DOMDocument();
		$this->doc_xml->encoding = mb_detect_encoding($xml);
		$this->doc_xml->preserveWhiteSpace = false;
		$this->doc_xml->formatOutput = true;
		$this->doc_xml->loadXML($xml);
		$this->doc_xml->saveXML();
		
		$this->rels_xml = new DOMDocument();
		$this->rels_xml->encoding = mb_detect_encoding($xml);
		$this->rels_xml->preserveWhiteSpace = false;
		$this->rels_xml->formatOutput = true;
		$this->rels_xml->loadXML($xml_rels);
		$this->rels_xml->saveXML();
		
	}


	/**
	 * 调用外部程序进行预处理替换原文档
	 */
	private function preprocessingWord($filename) {
		$fileDir = __DIR__.DIRECTORY_SEPARATOR.'b'.DIRECTORY_SEPARATOR.'c'.DIRECTORY_SEPARATOR;
		$exeName = __DIR__.DIRECTORY_SEPARATOR.'WordProcessing.exe';
		$exeName = 'WordProcessing.exe';
		$shell =  $exeName." $filename $fileDir";
		// echo $shell;die;
		$a = system("start E:\PHP\test\WordProcessing.exe E:\PHP\test\b\test.docx E:\PHP\test\b\c\\",$output);

		return $output;
	}


	/**
	 * CHECKS THE FONT FORMATTING OF A GIVEN ELEMENT
	 * Currently checks and formats: bold, italic, underline, background color and font family
	 * 
	 * @param XML $xml The XML node
	 * @return String HTML formatted code
	 */
	private function checkFormating(&$xml)
	{	
		// $node = trim($xml->readOuterXML());	
		// // echo  "<br/>" . $node . "*****************<br/>";
		// // add <br> tags
		// if (strstr($node,'<w:br ')) $text .= '<br>';					 
		// // look for formatting tags
		// $f = "<span style='";
		// $reader = new XMLReader();
		// $reader->XML($node);
		// while ($reader->read()) {
		// 	if($reader->name == "w:b")
		// 		$f .= "font-weight: bold,";
		// 	if($reader->name == "w:i")
		// 		$f .= "text-decoration: underline,";
		// 	if($reader->name == "w:color")
		// 		$f .="color: #".$reader->getAttribute("w:val").",";
		// 	if($reader->name == "w:rFont")
		// 		$f .="font-family: #".$reader->getAttribute("w:ascii").",";
		// 	if($reader->name == "w:shd" && $reader->getAttribute("w:val") != "clear" && $reader->getAttribute("w:fill") != "000000")
		// 		$f .="background-color: #".$reader->getAttribute("w:fill").",";
		// }
		// $f = rtrim($f, ',');
		// $f .= "'>";
		
		// return $f.htmlentities($xml->expand()->textContent)."</span>";

		// return $xml->expand();
		return $xml->expand()->textContent;
	}
	
	/**
	 * CHECKS THE ELEMENT FOR UL ELEMENTS
	 * Currently under development
	 * 
	 * @param XML $xml The XML node
	 * @return String HTML formatted code
	 */
	private function getListFormating(&$xml)
	{	
		$node = trim($xml->readOuterXML());
		
		$reader = new XMLReader();
		$reader->XML($node);
		$ret="";
		$close = "";
		// while ($reader->read()){
		// 	if($reader->name == "w:numPr" && $reader->nodeType == XMLReader::ELEMENT ) {
				
		// 	}
		// 	if($reader->name == "w:numId" && $reader->hasAttributes) {
		// 		switch($reader->getAttribute("w:val")) {
		// 			case 1:
		// 				$ret['open'] = "<ol><li>";
		// 				$ret['close'] = "</li></ol>";
		// 				break;
		// 			case 2:
		// 				$ret['open'] = "<ul><li>";
		// 				$ret['close'] = "</li></ul>";
		// 				break;
		// 		}
				
		// 	}
		// }
		return $ret;
	}
	
	/**
	 * CHECKS IF THERE IS AN IMAGE PRESENT
	 * Currently under development
	 * 
	 * @param XML $xml The XML node
	 * @return String HTML formatted code
	 */
	
	
	/**
	 * CHECKS IF ELEMENT IS AN HYPERLINK
	 *  
	 * @param XML $xml The XML node
	 * @return Array With HTML open and closing tag definition
	 */
	private function getHyperlink(&$xml)
	{
		$ret = array('open'=>'<ul>','close'=>'</ul>');
		$link ='';
		if($xml->hasAttributes) {
			$attribute = "";
			while($xml->moveToNextAttribute()) {
				if($xml->name == "r:id")
					$attribute = $xml->value;
			}
			
			if($attribute != "") {
				$reader = new XMLReader();
				$reader->XML($this->rels_xml->saveXML());
				
				while ($reader->read()) {
					if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name=='Relationship') {
						if($reader->getAttribute("Id") == $attribute) {
							$link = $reader->getAttribute('Target');
							break;
						}
					}
				}
			}
		}
		
		if($link != "") {
			$ret['open'] = "<a href='".$link."' target='_blank'>";
			$ret['close'] = "</a>";
		}
		
		return $ret;
	}
	
	/**
	 * READS THE GIVEN DOCX FILE INTO HTML FORMAT
	 *  
	 * @param String $filename The DOCX file name
	 * @return String With HTML code
	 */
	public function readDocument($filename) {
		
		$this->readZipPart($filename);
		$reader = new XMLReader();
		$reader->XML($this->doc_xml->saveXML());
		$text = ''; $list_format="";
		$front = '';
		$formatting['header'] = 0;
		$s = [];
		$bFlag = ['一、','二、','三、','四、','五、','六、','七、','八、','九、','十、'];
		$sFlag = 1;
		$flag = 0;
		$trFlag = 0;
		$tcFlag = 0;
		$teststr = '';
		$testoldstr = '';
		// $testoldstrArr = [];
		// $teststrArr = []; 
		// loop through docx xml dom
		while ($reader->read()) {
			$paragraph = new XMLReader;
			$p = $reader->readOuterXML();
			var_dump($p);die;
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:p') {
				$paragraph->xml($p);

				while ($paragraph->read()) {
					// var_dump($paragraph->name);
					if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:p') {
						$t = str_replace(array("\r\n", "\r", "\n"," "),'',$this->checkFormating($paragraph));
						if($t!=""){
							if(mb_substr($t,0,2, 'utf-8')==$bFlag[$flag]){
								$flag++;
								$sFlag = 2;
							} else if((mb_substr($t,0,2, 'utf-8')==='1.'||mb_substr($t,0,2, 'utf-8')==='1．')&&$sFlag!=2){
								$flag++;
								$s[$flag] .= $front . "*";
							} else if((mb_substr($t,0,2, 'utf-8')==='1.'||mb_substr($t,0,2, 'utf-8')==='1．')&&$sFlag==2) {
								$sFlag = 1;
							}
							if($flag == 0) {
								if(mb_strpos($t,'学年度')  !==  false||mb_strpos($t,'考试')  !==  false||mb_strpos($t,'学期')  !==  false||mb_strpos($t,'《')  !==  false||mb_strpos($t,'》')  !==  false){
									if(mb_strpos($t,'…') !==  false){
										$ex = explode('…',$t);
										$s[$flag] .= $ex[count($ex)-1];
									}
									else $s[$flag] .= $t . '*';
								}
							}
							else {
								if($t!='得分'&&$t!='评卷人')
									$s[$flag] .= $t . "*";
							}
							$front = $t;
							
						}
					}
					else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:t') { //list
						$list_format = $this->getListFormating($paragraph);
					}
					 if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:drawing') { //images

						// $a = $read->readInnerXml();
						// var_dump($a);







						// $a = new DOMDocument();
						// $a->preserveWhiteSpace = false;
						// $a->formatOutput = true;
						// $a->loadXML($paragraph);
						// $b = $a->childNodes->nodeName;
						// var_dump($b);die;
						// $b = $a->Xml;
						// $a = $paragraph->parentNode();
						// var_dump($a);
						// $text .= $this->checkImageFormating($paragraph);
					}
					else if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:hyperlink') {
						$hyperlink = $this->getHyperlink($paragraph);
						$text .= $hyperlink['open'];
						$text .= $this->checkFormating($paragraph);
						$text .= $hyperlink['close'];
					}
					else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:tbl') {
						// var_dump('this is table');
					}

					/**
					 * this here~ xiaosang read
					 */
					// $paragraph->next();
				}
			}
			else if($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:tbl'){
				$paragraph->xml($p);
				$tr_cache = '';
				$tr_int_boo=0;
				$teststr .= '&T<table>';
				while($paragraph->read()) {

					if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:tr') {
						$t = str_replace(array("\r\n", "\r", "\n"," "),'',$this->checkFormating($paragraph));
						if($t != "") {
							if(mb_substr($t,0,2, 'utf-8')==='得分' || mb_substr($t,0,2, 'utf-8')==='题号') {
								$teststr = '';
								continue;
							}else {
								$tr_cache = $t;
								$tr_int_boo = 0;
								$teststr .= '&!<tr>';
							}
						}
					}
					else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:tc') {
						$x = str_replace(array("\r\n", "\r", "\n"," "),'',$this->checkFormating($paragraph));
							if((mb_substr($x,0,2, 'utf-8')==='四、') && (mb_strpos($x,'(')  !==  false || mb_strpos($x,'（')  !==  false)) {
								$tcFlag = 1;
								continue;
							}else if(mb_substr($x,0,2, 'utf-8')==='得分' || mb_substr($x,0,3, 'utf-8')==='评卷人') {
								continue;
							}
							if($tcFlag == 1) {
								if(stristr($tr_cache,$x) !='' && $tr_int_boo == 0) {
									
									$tr_int_boo = 1;
								}
								if($tr_int_boo == 1) {
									$teststr .= "<tc>" . $x . "</tc>";
								}
							}
					}
				}
				if($teststr != '')
					$teststr .= '</tr></table>';
			}
			else if($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:drawing') {
				$paragraph->xml($p);
				while($paragraph->read()) {
					if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'a:blip') {
						$rId = $paragraph->getAttribute('r:embed');
						$rIdIndex = substr($rId,3);
						$imgStr = $this->checkImageFormating($rIdIndex);
						// var_dump($s);die;
						// $s[$flag] .= '<img>';
					}
				}

			}
		}
		$reader->close();
		$reArr = $this->handleStrTab($teststr);

		for($i = 1 ; $i < count($s); $i++) {
			for($j = 0; $j < count($reArr['old']); $j++) {
				if(stristr($s[$i],$reArr['old'][$j]) != '') {
					$s[$i] = str_replace($reArr['old'][$j],$reArr['new'][$j+1],$s[$i]);
				}
			}
		}
        return $this->division($s);
	}

	/**
	 * 不解压的情况下直接显示压缩包中的图片文件
	 */
	private function checkImageFormating($rIdIndex) {
		$imgname = 'word/media/image'.($rIdIndex-8);
		$zipfileName =  __DIR__.DIRECTORY_SEPARATOR.'b'.DIRECTORY_SEPARATOR.'test.docx';
		$zip=zip_open($zipfileName);
		while($zip_entry = zip_read($zip)) {//读依次读取包中的文件
			$file_name=zip_entry_name($zip_entry);//获取zip中的文件名    
			if(strstr($file_name,$imgname) != '') {
				if ($enter_zp = zip_entry_open($zip, $zip_entry, "r")) {  //读取包中文件
					$ext = pathinfo(zip_entry_name ($zip_entry),PATHINFO_EXTENSION);//获取图片文件扩展名
					$content = zip_entry_read($zip_entry,zip_entry_filesize($zip_entry));//读取文件二进制数据
					
					return  sprintf('<img src="data:image/%s;base64,%s">', $ext, base64_encode($content));//利用base64_encode函数转换读取到的二进制数据并输入输出到页面中
				}
				zip_entry_close($zip_entry); //关闭zip中打开的项目 
			}
		}
		zip_close($zip);//关闭zip文件    


		// $zip = new ZipArchive();
		// if (true === $zip->open($zipfileName)) {
		// 	if (($index = $zip->locateName($imgname)) !== false) {



		// 		$xml = $zip->getFromIndex($index);


		// 		// $path = './';
		// 		// $output_file = time().rand(100,999).'.gif';
		// 		// $path = $path.$output_file;
		// 		// $ifp = fopen( $path, "wb" );
		// 		// fwrite( $ifp, base64_decode($xml) );
		// 		// print_r($output_file);
		// 	}else {
		// 		var_dump('no file');
		// 	}
		// 	$zip->close();
		// } else die('non zip');
		
	}

    public function division ($s=[]) {
		$res = [];
		$res['title'] = mb_substr($s[0],0,mb_strpos($s[0],'考试')+2);
        $start = mb_strpos($s[0],'级');
        $end = mb_strpos($s[0],'专业');
        $res['course'] = mb_substr($s[0],$start+1,$end-$start-1);
        $res['question_types'] = [];

        for($i=0;$i<count($s)-1;$i++){
			$res['question_types'][$i] = [];
			$question = explode('*', $s[$i+1]);
			$qname = $question[0];
            if(mb_strpos($qname,'、')!==false){
				$qname= explode('、', $question[0])[1];
            } else {
                $qname = $question[0];
            }
            if(mb_strpos($qname,'(')!==false)
                $qname = explode('(', $qname)[0];
            else if(mb_strpos($qname,'（')!==false)
                $qname= explode('（', $qname)[0];
            $res['question_types'][$i]['name'] = $qname;
            $res['question_types'][$i]['title'] = $question[0];
            $res['question_types'][$i]['questions'] = [];

			$str = "";
            for($j=1;$j<count($question);$j++){
                $str .= $question[$j];
            }
			$flag = 2;
            while(true){
                if(mb_strpos($str,$flag . '．')!==false){
					if(mb_substr($str,mb_strpos($str,$flag . '．')-1,1) == '．') {
						$spl = mb_substr($str,0,mb_strpos($str,'．' . $flag . '．')).mb_substr($str,mb_strpos($str,'．'.$flag . '．'),mb_strpos($str,$flag . '．'));
						$res['question_types'][$i]['questions'][] = $spl;
						$str = str_replace($spl,'',$str);
					}else {
						$res['question_types'][$i]['questions'][] = mb_substr($str,0,mb_strpos($str,$flag . '．'));
						$str =  mb_substr($str,mb_strpos($str,$flag . '．'));
					}
                }else if(mb_strpos($str,$flag . '.')!==false){
					if(mb_substr($str,mb_strpos($str,$flag . '.')-1,1) == '.') {
						$spl = mb_substr($str,0,mb_strpos($str,'.' . $flag . '.')).mb_substr($str,mb_strpos($str,'.'.$flag . '.'),mb_strpos($str,$flag . '.'));
						$res['question_types'][$i]['questions'][] = $spl;
						$str = str_replace($spl,'',$str);
					}else {
						$res['question_types'][$i]['questions'][] = mb_substr($str,0,mb_strpos($str,$flag . '.'));
						$str =  mb_substr($str,mb_strpos($str,$flag . '.'));
					}
                }
                else break;
                $flag++;
			}
			if($str !== "") 
            	$res['question_types'][$i]['questions'][] = $str;
		}
		$this->setSelection($res);
		return $res;
	}
	
	public function setSelection(&$arr = []) {
		$titNam = '选择';
		$option = ['A','B','C','D'];
		for($a = 0; $a<count($arr);$a++) {
			$arr2 = &$arr['question_types'][$a];
			if(mb_strpos($arr2['name'],$titNam)!==false) {
				for ($i=0; $i < count($arr2['questions']); $i++) { 
					$str = $arr2['questions'][$i];
					$arr2['questions'][$i] = [];
					if(mb_strpos($str,$option[0] . '．')!==false) {
						$arr2['questions'][$i]['title'] = mb_substr($str,0,mb_strpos($str,$option[0] . '．'));
						$str =  mb_substr($str,mb_strpos($str,$option[0] . '．'));
					}
					for($j = 1; $j < count($option); $j++) {
						$arr2['questions'][$i]['options'][] =  mb_substr($str,0,mb_strpos($str,$option[$j] . '．'));
						$str =  mb_substr($str,mb_strpos($str,$option[$j] . '．'));
					}
					if($str !== "") $arr2['questions'][$i]['options'][] = $str;
				}
			}else return;
		}
	}

	public function handleStrTab($str = '') {
		$x = '';
		$arr = explode('&!',$str);
		for($i = 0; $i < count($arr); $i++) {
			if(substr($arr[$i],-5) == '</tc>') {
				$arr[$i] .= '</tr>';
			}
			$x .= $arr[$i];
		}
		$arr2Cache = $arr2 = explode('&T',$x);
		for($i = 1; $i < count($arr2Cache); $i++) {
			$str = '';
			$r = explode('<tc>',$arr2Cache[$i]);
			for($j = 0; $j < count($r); $j++) {
				$st = stristr($r[$j],"</tc>",true);
				if($st != '') {
					$str .= '*'.$st;
				}
			}
			$arr3[$i-1] = $str;
		}
		return ['new' => $arr2,'old' => $arr3];
	}
}