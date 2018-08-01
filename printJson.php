<?php
require './sang_cache.php';
$rt = new Word2Json();
$fileName = __DIR__.DS.'b'.DS.'test1.docx';
$res = $rt->readDocument($fileName);

echo '<pre>';
print_r(json_encode($res, JSON_UNESCAPED_UNICODE|JSON_PRETTY_PRINT)) ;

?>