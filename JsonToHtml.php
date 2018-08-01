<?php

// require(__DIR__.DIRECTORY_SEPARATOR.'libs'.DIRECTORY_SEPARATOR.'Smarty.class.php');



require './sang_cache.php';
$rt = new Word2Json();
$fileName = __DIR__.DS.'b'.DS.getUrl();
$res = $rt->readDocument($fileName);
$json2html = new JsonToHtml($res);

function getUrl() {
    $root = $_SERVER['SCRIPT_NAME'];
    $request = $_SERVER['REQUEST_URI'];
    $URI = array();
    $url = trim(str_replace($root, '', $request), '/');
    if(empty($url)) {
        var_dump('请添加路由参数');
        die;
    }
    else {
        $URI = explode('.', $url);
        if(count($URI) < 2) {
            $file = $URI[0] . '.docx';
        }
        else if($URI[1] !== 'docx') {
            var_dump('文件格式不正确');
            die;
        }
        else $file = $url;
    }
    return $file;
}


class JsonToHtml 
{

    private $json;
    private $html = '';
    public function __construct($val) {
        $this->json = is_null(json_decode($val)) ? json_decode(json_encode($val)) : $this->$val;
        if(is_object($this->json)) $this->json2html($this->json);

      
        
    }
    
    private function json2html(&$obj) {
        // var_dump($obj);die;
        $this->returnHeader();
        $this->html .= sprintf('<div style="%s">%s</div>', $obj->title->style, $obj->title->value);
        // $this->html .= '<div>'.$obj->title->value.'</div>';

        $this->returnTable(count($obj->question_types), $obj->title->value);
        $this->returnContent($obj->question_types);
        echo $this->html;
        

        
    }

    private function returnHeader() {
        $this->html .= '<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <meta http-equiv="X-UA-Compatible" content="ie=edge">
            <title>Document</title>
        </head>
        <body>';
    }

    private function returnTable($num, $str) {
        $tblFlag = ['一','二','三','四','五','六','七','八','九','十'];
        $this->html .= sprintf('<div><table border="1" style="display:table;margin:0 auto;border-collapse:collapse;text-align:center;width:%s">', $this->getTitleWidth($str));
        for($i = 0; $i < 2; $i++) {
            $this->html .= '<tr>';
            for($j = 0; $j < $num; $j++) {
                if($i == 0) {
                    if($j == 0) $this->html .= '<th>题号</th>';
                    $this->html .= '<th>'.$tblFlag[$j].'</th>';
                    if($j == $num - 1) $this->html .= '<th>合分</th><th>合分人</th><th>复核人</th></tr>';
                }
                else {
                    if($j == 0) $this->html .= '<td>得分</td>';
                    $this->html .= '<td></td>';
                    if($j == $num -1) $this->html .= '<td></td><td></td><td></td></tr>';
                }
            }
        }
        $this->html .= '</table></div>';
    }

    private function getTitleWidth($str) {
        $str = explode('<br>',$str)[0];
        $len = mb_strlen($str,"utf-8");
        
        for($i = 0 ; $i < $len ; $i++) {
            $a = mb_substr($str, 0, 1);
            if(mb_detect_encoding($a) == 'UTF-8') $zh += 1;
            $str = mb_substr($str, 1);
        }
        $pt_len = ($zh * 16) + ((($len-$zh) / 2) * 16); 
        return $pt_len . 'pt';
    }


    private function returnContent($questions) {
        $this->html .= '<div class="content">';
        foreach($questions as $key => $value) {
            // var_dump($value);die;
            if(strstr($value->name->value,'选择')) $this->returnChoice($value);
            elseif(strstr($value->name->value,'判断')) $this->returnJudge($value);
            elseif(strstr($value->name->value,'简答') || strstr($value->name->value,'解答')) $this->returnBrief($value);
            elseif(strstr($value->name->value,'计算') || strstr($value->name->value,'证明')) $this->returnCalculation($value);
            elseif(strstr($value->name->value,'填空')) $this->returnPack($value);

        }
        $this->html .= '</div>';
    }


    private function returnChoice($value) {
        $this->html .= '<div class="Choice">';
        // $this->html .= '<h2>'.$value->title->value.'</h2>';
        $this->html .= sprintf('<p style="%s">%s</p>',$value->title->style,$value->title->value);
        for($i = 0; $i <count($value->questions); $i++) {
            $this->html .= '<div class="'.Choice.$i.'">';
            $this->html .= sprintf('<p style="%s">%s</p>',$value->questions[$i]->title->style,$value->questions[$i]->title->value); 
            for($j = 0; $j < count($value->questions[$i]->options);$j++) {
                $this->html .= sprintf('<p style="%s">%s</p>',$value->questions[$i]->options[$j]->style,$value->questions[$i]->options[$j]->value);  
            }
            $this->html .= '</div>';
        }
        $this->html .= '</div>';
    }

    private function returnJudge($value) {
        $this->html .= '<div class="Judge">';
        $this->html .= sprintf('<p style="%s">%s</p>',$value->title->style,$value->title->value);
        for($i = 0; $i < count($value->questions); $i++) {
            $this->html .= sprintf('<p style="%s">%s</p>',$value->questions[$i]->style,$value->questions[$i]->value);   
        }
        $this->html .= '</div>';
    }

    private function returnBrief($value) {
        $this->html .= '<div class="Brief">';
        $this->html .= sprintf('<p style="%s">%s</p>',$value->title->style,$value->title->value);
        for($i = 0; $i < count($value->questions); $i++) {
            if(is_string($value->questions[$i]->value))
                $this->html .= sprintf('<p style="%s">%s</p>',$value->questions[$i]->style,$value->questions[$i]->value);  
            else if(is_object($value->questions[$i]->value)) {
                $this->html .= sprintf('<div><p style="%s">%s</p>',$value->questions[$i]->value->secondsTitle->style,$value->questions[$i]->value->secondsTitle->value);  
                foreach ($value->questions[$i]->value->subtitle as $key => $item) {
                    $this->html .= sprintf('<p style="%s">%s</p>',$item->style,$item->value); 
                }
                $this->html .= '</div>';
            }
        }
        $this->html .= '</div>';

    }

    private function returnCalculation($value) {
        // var_dump($value);die;
        $this->html .= '<div class="Calculation">';
        $this->html .= sprintf('<p style="%s">%s</p>', $value->title->style, $value->title->value); 
        for($i = 0; $i < count($value->questions); $i++) {
            if(is_string($value->questions[$i]->value))
                $this->html .= sprintf('<div style="%s">%s</div>', $value->questions[$i]->style, $value->questions[$i]->value);
            else if(is_object($value->questions[$i]->value)) {
                $this->html .= sprintf('<div><p style="%s">%s</p>', $value->questions[$i]->value->secondsTitle->style, $value->questions[$i]->value->secondsTitle->value);
                foreach ($value->questions[$i]->value->subtitle as $key => $item) {
                    $this->html .= sprintf('<p style="%s">%s</p>', $item->style, $item->value);
                }
                $this->html .= '</div>';
            }
        }
        $this->html .= '</div>';
    }

    private function returnPack($value) {
        $this->html .= '<div class="Pack">';
        $this->html .= sprintf('<p style="%s">%s</p>',$value->title->style,$value->title->value);
        for($i = 0; $i < count($value->questions); $i++) {
            $this->html .= sprintf('<p style="%s">%s</p>', $value->questions[$i]->style, $value->questions[$i]->value); 
        }
        $this->html .= '</div>';
    }


}

