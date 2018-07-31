<?php

define('DS') || define('DS',DIRECTORY_SEPARATOR);

function __autoload($className) {
	require_once __DIR__.DS.$className.'.php';
}


$rt = new Word2Json();

$fileName = __DIR__.DS.'b'.DS.'test1.docx';
$res = $rt->readDocument($fileName);
$json2html = new JsonToHtml($res);

class Word2Json
{
	private $rels_xml;
	private $doc_xml;
	

	/**
	 * 判断并将word转换为xml
	 */
	private function readZipPart($filename) {
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
		$fileDir = __DIR__.DS.'docx'.DS.'aaa'.DS;
		$exeName = 'WordProcessing.exe';
		$shell =  $exeName." $filename $fileDir";
		$a = exec($shell,$output,$return_val);
		return $return_val;
	}


	private function checkFormating(&$xml) {	
		return $xml->expand()->textContent;
	}
	
	private function getListFormating(&$xml) {	
		$node = trim($xml->readOuterXML());
		
		$reader = new XMLReader();
		$reader->XML($node);
		$ret="";
		$close = "";
		return $ret;
	}
	

	/**
	 * 处理文档中的链接
	 */
	private function getHyperlink(&$xml) {
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
	

	//主要实现
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

		while ($reader->read()) {
			$paragraph = new XMLReader;
			$p = $reader->readOuterXML();
			// var_dump($p);die;
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:p') {
				$paragraph->xml($p);
				while ($paragraph->read()) {
					if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:p') {
						$t ='';
						$cache_t = [];
						$i = 0;
						while($paragraph->read()) {
							if($paragraph->nodeType == XMLReader::ELEMENT) { 
								$ts = str_replace(array("\r\n", "\r", "\n"," "),'',$this->checkFormating($paragraph));

								//判断是否为空格
								if(($paragraph->name === 'w:r') 
								&& (mb_strpos($paragraph->readOuterXml(),'<w:u w:val="single"/>') >-1) 
								&& (mb_strpos($paragraph->readOuterXml(),'<w:t xml:space="preserve">') > -1)) {
									$underlineLength = mb_strpos($paragraph->readOuterXml(),'</w:t>') - (mb_strpos($paragraph->readOuterXml(),'<w:t xml:space="preserve">') + strlen('<w:t xml:space="preserve">'));
									$t .= $this->test($underlineLength);
								}
								

								//检测是否为编号
								if(($paragraph->name === 'w:numPr') 
								&& (mb_strpos($paragraph->readOuterXml(), 'w:ilvl') > -1) 
								&& (mb_strpos($paragraph->readOuterXml(), 'w:numId') > -1)) {
									if(($lastVal = mb_substr($paragraph->readOuterXml(), mb_strripos($paragraph->readOuterXml(), 'w:val') + 7, 1)) 
									&& is_numeric($lastVal) 
									&& is_numeric($firstVal = mb_substr($paragraph->readOuterXml(), mb_strpos($paragraph->readOuterXml(), 'w:val') + 7, 1))) {
										$t .= '&(';
									}
								}

				
								
								if($ts == '' && $paragraph->name != 'w:drawing') continue;
								if($paragraph->name === 'w:drawing') {
									// (strstr($ts,'…封…') != false || strstr($ts,'…线…') != false) ? $t .= '' : $t .= $this->analysisDrawing($paragraph);
									(strstr($ts,'…封…') != false || strstr($ts,'…线…') != false) ? $t .= '' : $t .= '<img>';
								}
								if((strstr($t,$ts) == false) || (strstr($cache_t[count($cache_t)-1],$ts) ==false)){
									if(strstr($cache_t[count($cache_t)-1],$ts) === '0') continue;
									$t .= $ts;
								}else continue;
								$cache_t[$i++] = $ts;
							}
						}
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
					else if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:hyperlink') {
						$hyperlink = $this->getHyperlink($paragraph);
						$text .= $hyperlink['open'];
						$text .= $this->checkFormating($paragraph);
						$text .= $hyperlink['close'];
					}
				}
			}
			else if($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:tbl'){
				$paragraph->xml($p);
				$tr_cache = '';
				$tr_int_boo=0;
				$teststr .= '&T<table border="1">';
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
								$teststr .= "<th>" . $x . "</th>";
							}
						}
					}
				}
				if($teststr != '')
					$teststr .= '</tr></table>';
			}
		}
		$reader->close();
		$reArr = $this->handleStrTab($teststr);
		for($i = 1 ; $i < count($s); $i++) 
			for($j = 0; $j < count($reArr['old']); $j++)
				if(stristr($s[$i],$reArr['old'][$j]) != '') 
					$s[$i] = str_replace($reArr['old'][$j],$reArr['new'][$j+1],$s[$i]);
		

        return $this->division($s);
	}


	/**
	 * 获取图片索引值
	 */
	private function analysisDrawing(&$drawingXml) {
		$rIdIndex = '';
		$distArr = [];
		$slideSizeArr = [];
		while($drawingXml->read())
			if($drawingXml->nodeType == XMLREADER::ELEMENT) 
				switch($drawingXml->name)
				{
					case 'wp:inline' : 
						$distName = ['distT', 'distB', 'distL', 'distR'];
						foreach ($distName as $dist)
							$distArr[$dist] = (!is_numeric($drawingXml->getAttribute($dist))) ? :$this->emuToPx(intval($drawingXml->getAttribute($dist)));
						break;
					case 'wp:extent' :
						$slideSizeArr['cx'] = (!is_numeric($drawingXml->getAttribute('cx'))) ? :$this->emuToPx(intval($drawingXml->getAttribute('cx')));
						$slideSizeArr['cy'] = (!is_numeric($drawingXml->getAttribute('cy'))) ? :$this->emuToPx(intval($drawingXml->getAttribute('cy')));
						break;
					case 'a:blip' : 
						$rId = $drawingXml->getAttribute('r:embed');
						$rIdIndex = substr($rId,3);
						return $this->checkImageFormating($rIdIndex, $distArr, $slideSizeArr);
						break;
				}
	}

	/**
	 * 找到并读取图片流，转化为base64格式进行拼接显示图片
	 */
	private function checkImageFormating($rIdIndex, $distArr = [], $slideSizeArr = []) {

		$imgname = 'word/media/image'.($rIdIndex-8);
		$zipfileName =  __DIR__.DS.'b'.DS.'test.docx';
		$zip=zip_open($zipfileName);
		while($zip_entry = zip_read($zip)) {//读依次读取包中的文件
			$file_name=zip_entry_name($zip_entry);//获取zip中的文件名
			if(strstr($file_name,$imgname) != '' ) {
				$a = ($rIdIndex-8 < 10) ? mb_substr($file_name,mb_strlen($imgname,"utf-8"),1, 'utf-8') : '';    
				if($rIdIndex-8 < 10 && $a != '.') continue;
				if ($enter_zp = zip_entry_open($zip, $zip_entry, "r")) {  //读取包中文件
					$ext = pathinfo(zip_entry_name ($zip_entry),PATHINFO_EXTENSION);//获取图片文件扩展名
					$content = zip_entry_read($zip_entry,zip_entry_filesize($zip_entry));//读取文件二进制数据
					// return sprintf('<img src="data:image/%s;base64,%s" style="'.'width:'.$slideSizeArr['cx'].';height:'.$slideSizeArr['cy'].';">', $ext, base64_encode($content));//利用base64_encode函数转换读取到的二进制数据并输入输出到页面中
					return sprintf('<img src="data:image/%s;base64,%s" style="width:%s;height:%s;margin-top:%s;margin-bottom:%s;margin-left:%s;margin-right:%s">', $ext, base64_encode($content), ...array_values($slideSizeArr), ...array_values($distArr));//利用base64_encode函数转换读取到的二进制数据并输入输出到页面中
				}
				zip_entry_close($zip_entry); //关闭zip中打开的项目 
			}
		}
		zip_close($zip);//关闭zip文件   
	}

    public function division ($s=[]) {
		// var_dump($s);die;
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
		var_dump($res);die;
		$this->setSelection($res);

		$this->checkSubtitle($res['question_types']);
		return $res;
	}
	

	/**
	 * 分割选择题选项
	 */
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
					}else if(mb_strpos($str,$option[0] . ' ．')!==false) {
						$arr2['questions'][$i]['title'] = mb_substr($str,0,mb_strpos($str,$option[0] . ' ．'));
						$str =  mb_substr($str,mb_strpos($str,$option[0] . ' ．'));
					}
					for($j = 1; $j < count($option); $j++) {
						if(mb_strpos($str,$option[$j] . '．')) {
							$arr2['questions'][$i]['options'][] =  mb_substr($str,0,mb_strpos($str,$option[$j] . '．'));
							$str =  mb_substr($str,mb_strpos($str,$option[$j] . '．'));
						}
						else if(mb_strpos($str,$option[$j] . ' ．')) {
							$arr2['questions'][$i]['options'][] =  mb_substr($str,0,mb_strpos($str,$option[$j] . ' ．'));
							$str =  mb_substr($str,mb_strpos($str,$option[$j] . ' ．'));
						}
					}
					if($str !== "") $arr2['questions'][$i]['options'][] = $str;
				}
			}else return;
		}
	}


	/**
	 * 检查题中是否有序列号标识，若有，替换成序列号
	 */
	public function checkSubtitle(&$topicBankArr = []) {
		for ($i = 2; $i < count($topicBankArr); $i++) { 
			$arr = &$topicBankArr[$i]['questions'];
			for ($j = 0; $j < count($arr); $j++) {
				if((mb_strpos($arr[$j], '&(') > -1) || (mb_strpos($arr[$j], '&（') > -1)) {
					$str = '';
					$sequence = 0;
					(count(explode('&(',$arr[$j])) > 1) ? ($arr2 = explode('&(',$arr[$j])) : ($arr2 = explode('&（',$arr[$j]));
					foreach ($arr2 as $key => $value) 
						($key == 0) ? ($str .= $value) : ($str .= ('(' . ++$sequence . ')' . $value));
					$arr[$j] = $str;
				}
				$arr[$j] = $this->divisionSubtitle($arr[$j]);
			}
		}
	}

	/**
	 * 分割子标题
	 */
	public function divisionSubtitle($str = '') {
		$tit = [1 ,2, 3, 4, 5, 6];
		for ($i = 0; $i < count($tit); $i++) { 
			if($i == 0 && !(substr_count($str, '(' . $tit[$i] . ')') || substr_count($str, '（' . $tit[$i] . '）')))
				return $str;
			
			if(!substr_count($str, '(' . $tit[$i] . ')') && !substr_count($str, '（' . $tit[$i] . '）')) {
				$arr['subtitle'][$i -1] = $str;
				break;
			}else {
				$subStr = mb_substr($str, 0, (mb_strpos($str, '(' . $tit[$i] . ')') | mb_strpos($str, '（' . $tit[$i] . '）')));
				substr_count($str, '(' . $tit[$i] . ')') ? ($str = stristr($str, '(' . $tit[$i] . ')')) : ($str = stristr($str, '（' . $tit[$i] . '）'));
				($i == 0) ? ($arr['secondsTitle'] = $subStr) : ($arr['subtitle'][$i -1] = $subStr);
			}
		}
		return $arr;
	}



	/**
	 * 拼接成表格
	 */
	public function handleStrTab($str = '') {
		$x = '';
		$arr = explode('&!',$str);
		for($i = 0; $i < count($arr); $i++) {
			if(substr($arr[$i],-5) == '</th>') {
				$arr[$i] .= '</tr>';
			}
			$x .= $arr[$i];
		}
		$arr2Cache = $arr2 = explode('&T',$x);
		for($i = 1; $i < count($arr2Cache); $i++) {
			$str = '';
			$r = explode('<th>',$arr2Cache[$i]);
			for($j = 0; $j < count($r); $j++) {
				$st = stristr($r[$j],"</th>",true);
				if($st != '') {
					$str .= '*'.$st;
				}
			}
			$arr3[$i-1] = $str;
		}
		return ['new' => $arr2,'old' => $arr3];
	}



	/**
	 *  英语公制单位转磅
	 * Emu  ->  pt
	 */
	private function emuToPx($emu = 0) {
		$e = 914400;
		$pt = 72;
		return  round(($emu / $e) * $pt, 1) . 'pt';
	}



	/**
	 * test echo 
	 */
	public function echos($a,$b,$c) {
		echo $a.'#'.$b.'#'.$c;
	}


	/**
	 * 生成填空题下滑线
	 */
	public function test($ints) {
		$a = '';
		for($i = 0; $i < $ints; $i++) $a .= '_';
		return $a;
	}
}