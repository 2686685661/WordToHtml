<?php

// require(__DIR__.DIRECTORY_SEPARATOR.'libs'.DIRECTORY_SEPARATOR.'Smarty.class.php');

class JsonToHtml 
{

    private $json;
    private $html = '';
    public function __construct($val) {
        $this->json = is_null(json_decode($val)) ? json_decode(json_encode($val)) : $this->$val;
        if(is_object($this->json)) $this->json2html($this->json);
        
    }
    
    private function json2html(&$obj) {
        $this->returnHeader();
        $this->html .= '<div>'.$obj->title.'</div>';
        $this->returnTable(count($obj->question_types));
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

    private function returnTable($num) {
        $tblFlag = ['一','二','三','四','五','六','七','八','九','十'];
        $this->html .= '<table border="1">';
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
        $this->html .= '</table>';
    }


    private function returnContent($questions) {
        $this->html .= '<div class="content">';
        foreach($questions as $key => $value) {
            if(strstr($value->name,'选择')) $this->returnChoice($value);
            elseif(strstr($value->name,'判断')) $this->returnJudge($value);
            elseif(strstr($value->name,'简答') || strstr($value->name,'解答')) $this->returnBrief($value);
            elseif(strstr($value->name,'计算') || strstr($value->name,'证明')) $this->returnCalculation($value);
            elseif(strstr($value->name,'填空')) $this->returnPack($value);

        }
        $this->html .= '</div>';
    }


    private function returnChoice($value) {
        $this->html .= '<div class="Choice">';
        $this->html .= '<h2>'.$value->title.'</h2>';
        for($i = 0; $i <count($value->questions); $i++) {
            $this->html .= '<div class="'.Choice.$i.'">';
            $this->html .= '<p>'.$value->questions[$i]->title.'</p>';
            for($j = 0; $j < count($value->questions[$i]->options);$j++) {
                $this->html .= '<p>'.$value->questions[$i]->options[$j].'</p>';
            }
            $this->html .= '</div>';
        }
        $this->html .= '</div>';
    }

    private function returnJudge($value) {
        $this->html .= '<div class="Judge">';
        $this->html .= '<h2>'.$value->title.'</h2>';
        for($i = 0; $i < count($value->questions); $i++) {
            $this->html .= '<p>'.$value->questions[$i].'</p>';
        }
        $this->html .= '</div>';
    }

    private function returnBrief($value) {
        $this->html .= '<div class="Brief">';
        $this->html .= '<h2>'.$value->title.'</h2>';
        for($i = 0; $i < count($value->questions); $i++) {
            $this->html .= '<p>'.$value->questions[$i].'</p>';
        }
        $this->html .= '</div>';

    }

    private function returnCalculation($value) {
        $this->html .= '<div class="Calculation">';
        $this->html .= '<h2>'.$value->title.'</h2>';
        for($i = 0; $i < count($value->questions); $i++) {
            $this->html .= '<div>'.$value->questions[$i].'</div>';
        }
        $this->html .= '</div>';
    }

    private function returnPack($value) {
        $this->html .= '<div class="Pack">';
        $this->html .= '<h2>'.$value->title.'</h2>';
        for($i = 0; $i < count($value->questions); $i++) {
            $this->html .= '<p>'.$value->questions[$i].'</p>';
        }
        $this->html .= '</div>';
    }


}

