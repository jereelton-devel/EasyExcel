<?php

error_reporting(E_ALL);
ini_set('display_errors', true);
ini_set('display_startup_errors', true);

echo "<pre>";
echo "<br />[INICIO]<br />";

//Libs
$easyexcel = "./lib/sistema/EasyExcel.php";
$phpexcel  = "./lib/sistema/PHPExcel/Classes/PHPExcel/IOFactory.php";

//Propiedades do documento
$props = array('Jereelton','Jereelton','Teste','Teste Subject','Teste Description','Teste Keywords','Teste Category');

//Header da planilha
$header = array('Nome Campo 1','Nome Campo 2','Nome Campo 3','Nome Campo 4','Formula','Total','Formula','Media');

//Estilos e Formatações da planilha
$styles = array(
    'hidden_column' => 'E,G',
    'freeze_row' => 'A2',
    'freeze_byrow' => '2',
    'freeze_bycolumn' => '1',
    'autosize_all' => '1',
    'autosize_uniq' => '',
    'autosize_range_ini' => '',
    'autosize_range_fin' => '',
    'cells_align' => 'A1:H2560',
    'horizontal_align' => 'center',
    'vertical_align' => 'center',
    'fonte_color' => array(
        0 => [
            'cells_range' => 'A1:A1',
            'cells_color' => 'yellow'
        ],
        1 => [
            'cells_range' => 'B1:D1',
            'cells_color' => 'green'
        ],
        2 => [
            'cells_range' => 'E1:H1',
            'cells_color' => 'FFFFFFFF'
        ]
    ),
    'fill_color' => array(
        0 => [
            'cells_range' => 'A1:A1',
            'cells_color' => 'black'
        ],
        1 => [
            'cells_range' => 'B1:D1',
            'cells_color' => 'blue'
        ],
        2 => [
            'cells_range' => 'E1:H1',
            'cells_color' => 'red'
        ]
    ),
    'filter' => 'A1:H1'
);

//Dados da planilha
$param = array(
    0 => [
        'NomeCampo1' => 'Valor do campo 1 com largura maior',
        'NomeCampo2' => 20,
        'NomeCampo3' => 10,
        'NomeCampo4' => 15
    ],
    1 => [
        'NomeCampo1' => 'Valor do campo 1 com largura maior',
        'NomeCampo2' => 12,
        'NomeCampo3' => 4,
        'NomeCampo4' => 48
    ],
    2 => [
        'NomeCampo1' => 'Valor do campo 1 com largura maior',
        'NomeCampo2' => 8,
        'NomeCampo3' => 12,
        'NomeCampo4' => 5
    ],
    3 => [
        'NomeCampo1' => 'Valor do campo 1 com largura maior',
        'NomeCampo2' => 50,
        'NomeCampo3' => 30,
        'NomeCampo4' => 19
    ]
);

//Operacoes Matematicas
$oper = array(
    0 => ['oper' => '=SUM', 'start' => 'B', 'end' => 'D', 'ftarget' => 'E', 'rtarget' => 'F', 'mode' => 'repeat', 'char' => ':'],
    1 => ['oper' => '=AVERAGE', 'start' => 'B', 'end' => 'D', 'ftarget' => 'G', 'rtarget' => 'H', 'mode' => 'repeat', 'char' => ':']
);

//Nome do documento
$fname = "Arquivo_Teste_".date('dmYHis').".xlsx";

//Importa a lib para trabalhar com PHPExcel
require_once($easyexcel);

$easyExcelObj = new EasyExcel($fname, $phpexcel);echo "<br />[EasyExcel]<br />";
//$param = $easyExcelObj->getDataSample();echo "<br />[getDataSample]<br />";
//$oper = $easyExcelObj->getOperationsSample();echo "<br />[getOperationsSample]<br />";
$easyExcelObj->setFileProps($props);echo "<br />[setFileProps]<br />";
$easyExcelObj->setXlsHeader($header);echo "<br />[setXlsHeader]<br />";
$easyExcelObj->setXlsContent(2, $param, $oper);echo "<br />[setXlsContent]<br />";
$easyExcelObj->setXlsStyles($styles);echo "<br />[setXlsStyles]<br />";
$easyExcelObj->saveAs('excelObj');echo "<br />[saveAs]<br />";

echo "Arquivo: ". $fname;
echo "<br />[FIM]<br />";
echo "</pre>";

?>

