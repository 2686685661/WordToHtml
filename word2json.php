<?php

/**
 * time:no time
 * auther:小桑
 * function:初版，初步转为json,存在bug
 */
$rt = new Word2Json();

$res = $rt->readDocument('test1.docx');

var_dump($res); 


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
	private function checkImageFormating(&$xml) {
		
	}
	
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
		// loop through docx xml dom
		while ($reader->read()) {

		// 	// look for new paragraphs
			$paragraph = new XMLReader;
			// $reader->read();
			$p = $reader->readOuterXML();
			// echo "<textarea>" . $p . "</textarea>";
			// break;
			// echo $reader->nodeType . ')' . $reader->name . "<br/>";
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name === 'w:p') {
				// set up new instance of XMLReader for parsing paragraph independantly				
				$paragraph->xml($p);
				// preg_match('/<w:pStyle w:val="(Heading.*?[1-6])"/',$p,$matches);
				// if(isset($matches[1])) {
				// 	switch($matches[1]){
				// 		case 'Heading1': $formatting['header'] = 1; break;
				// 		case 'Heading2': $formatting['header'] = 2; break;
				// 		case 'Heading3': $formatting['header'] = 3; break;
				// 		case 'Heading4': $formatting['header'] = 4; break;
				// 		case 'Heading5': $formatting['header'] = 5; break;
				// 		case 'Heading6': $formatting['header'] = 6; break;
				// 		default: $formatting['header'] = 0; break;
				// 	}
				// }
				// // open h-tag or paragraph
				// $text .= ($formatting['header'] > 0) ? '<h'.$formatting['header'].'>' : '<p>';
				
				// loop through paragraph dom
				while ($paragraph->read()) {
					if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:p') {
						// if($list_format == "")
						// 	$text .= $this->checkFormating($paragraph);
						// else {
						// 	$text .= $list_format['open'];
						// 	$text .= $this->checkFormating($paragraph);
						// 	$text .= $list_format['close'];
						// }
						// $list_format ="";
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
					else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:pPr') { // 
						$list_format = $this->getListFormating($paragraph);
					}
					else if($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:drawing') { //images
						$text .= $this->checkImageFormating($paragraph);
					}
					else if ($paragraph->nodeType == XMLREADER::ELEMENT && $paragraph->name === 'w:hyperlink') {
						$hyperlink = $this->getHyperlink($paragraph);
						$text .= $hyperlink['open'];
						$text .= $this->checkFormating($paragraph);
						$text .= $hyperlink['close'];
					}
					$paragraph->next();
				}
			}
		}
		$reader->close();
		// var_dump($s);die;
        return $this->division($s);
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
            // var_dump($question);die;
			$qname = $question[0];
            if(mb_strpos($qname,'、')!==false){
				$qname= explode('、', $question[0])[1];
            } else {
                $qname = $question[0];
                // $res['question_types'][$i]['name']= explode('(', $qname)[0];
            }
            if(mb_strpos($qname,'(')!==false)
                $qname = explode('(', $qname)[0];
            else if(mb_strpos($qname,'（')!==false)
                $qname= explode('（', $qname)[0];
            $res['question_types'][$i]['name'] = $qname;
            $res['question_types'][$i]['title'] = $question[0];
            $res['question_types'][$i]['questions'] = [];

			$str = "";
			// var_dump($question);die;
            for($j=1;$j<count($question);$j++){
                $str .= $question[$j];
            }
			$flag = 2;
			// echo $str;die;
            while(true){
                if(mb_strpos($str,$flag . '．')!==false){
					// echo $flag . '．';
					$res['question_types'][$i]['questions'][] = mb_substr($str,0,mb_strpos($str,$flag . '．'));
                    $str =  mb_substr($str,mb_strpos($str,$flag . '．'));
                }else if(mb_strpos($str,$flag . '.')!==false){
                    $res['question_types'][$i]['questions'][] = mb_substr($str,0,mb_strpos($str,$flag . '.'));
                    $str =  mb_substr($str,mb_strpos($str,$flag . '.'));
                }
                else break;
                $flag++;
			}
			// var_dump($str);die;
            $res['question_types'][$i]['questions'][] = $str;
		}
        return $res;
    }
}