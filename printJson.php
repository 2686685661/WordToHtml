<?php
require './sang_cache.php';
$rt = new Word2Json();
$fileName = __DIR__.DS.'b'.DS.'test.docx';
$res = $rt->readDocument($fileName);
var_dump($res);

?>