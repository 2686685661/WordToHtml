<?php
require './sang_cache.php';
$rt = new Word2Json();
$fileName = __DIR__.DS.'b'.DS.getUrl();
$res = $rt->readDocument($fileName);

echo '<pre>';
print_r(json_encode($res, JSON_UNESCAPED_UNICODE|JSON_PRETTY_PRINT)) ;


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

?>