<?php
declare(strict_types=1);
$root=dirname(__DIR__);
$core=dirname($root).'/mnb-phpexcel-core/src/';
spl_autoload_register(static function(string $class) use($root,$core): void {
    $prefix='Mnb\\PHPExcel\\'; if(!str_starts_with($class,$prefix)) return;
    $rel=str_replace('\\',DIRECTORY_SEPARATOR,substr($class,strlen($prefix))).'.php';
    foreach([$root.'/src/'.$rel,$core.$rel] as $p) if(is_file($p)) { require $p; return; }
});
use Mnb\PHPExcel\Format\Csv;
$file=tempnam(sys_get_temp_dir(),'mnbmeta_').'.csv';
file_put_contents($file, "name,amount\nA,10\n");
try {
 $m=Csv::meta($file)->quick()->only(['file','workbook'])->read();
 if($m->format()!=='csv') throw new RuntimeException('format');
 if(!isset($m['file'])) throw new RuntimeException('array access');
 echo "Csv MetadataFacade V2 passed.\n";
} finally { @unlink($file); }
