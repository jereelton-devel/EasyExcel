<?php

error_reporting(E_ALL);
ini_set('display_errors', true);
ini_set('display_startup_errors', true);

echo "<pre>";
echo "<br />[INICIO]<br />";

//Importa a lib para trabalhar com PHPExcel
include("./lib/sistema/PHPExcel/Classes/PHPExcel/IOFactory.php");

//Nome do arquivo a ser gerado
$filename = 'arquivo-' . date('dmYHis') . '.xlsx';
echo "<br />{$filename}<br />";

//Dados a serem gravados no arquivo XLSX
$results = array(
    0 => [
        'NomeCampo1' => 'Valor do campo 1 com largura maior',
        'NomeCampo2' => 'Valor2',
        'NomeCampo3' => 10,
        'NomeCampo4' => 15
    ],
    1 => [
        'NomeCampo1' => 'Valor do campo 1 com largura maior',
        'NomeCampo2' => 'Valor2',
        'NomeCampo3' => 4,
        'NomeCampo4' => 48
    ],
    2 => [
        'NomeCampo1' => 'Valor do campo 1 com largura maior',
        'NomeCampo2' => 'Valor2',
        'NomeCampo3' => 12,
        'NomeCampo4' => 5
    ],
    3 => [
        'NomeCampo1' => 'Valor do campo 1 com largura maior',
        'NomeCampo2' => 'Valor2',
        'NomeCampo3' => 30,
        'NomeCampo4' => 19
    ]);

$operations = array(
    0 => [
        '=SUM' => [
            'start' => 'C',
            'end' => 'D',
            'mode' => 'repeat'
        ],
        '=AVERAGE' => [
            'start' => 'C',
            'end' => 'D',
            'mode' => 'repeat'
        ],
    ]);

var_dump($results);

//Instancia da classe Excel
$excelObj = new PHPExcel();

//var_dump($excelObj);

//Propriedades do documento
$excelObj->getProperties()
    ->setCreator("Websolutions Care")
    ->setLastModifiedBy("Jereelton Teixeira")
    ->setTitle("Titulo do Arquivo")
    ->setSubject("Titulo do Arquivo")
    ->setDescription("Dados exportados do Arquivo")
    ->setKeywords("chaves do arquivo")
    ->setCategory("categoria do arquivo");

//Ativa a primeira folha folha da planilha (indice 0)
$sheet = $excelObj->setActiveSheetIndex(0);

//Criando o cabeçalho da planilha
$sheet
    ->setCellValue('A1', 'Nome Campo 1')
    ->setCellValue('B1', 'Nome Campo 2')
    ->setCellValue('C1', 'Nome Campo 3')
    ->setCellValue('D1', 'Nome Campo 4')
    ->setCellValue('E1', 'Formula')
    ->setCellValue('F1', 'Total');

//Criando o conteudo da planilha
$cel_idx = 2;
for($i = 0; $i < count($results); $i++) {
    $sheet
        ->setCellValue('A'.$cel_idx, $results[$i]['NomeCampo1'])
        ->setCellValue('B'.$cel_idx, $results[$i]['NomeCampo2'])
        ->setCellValue('C'.$cel_idx, $results[$i]['NomeCampo3'])
        ->setCellValue('D'.$cel_idx, $results[$i]['NomeCampo4'])
        ->setCellValue('E'.$cel_idx, '=SUM(C'.$cel_idx.':D'.$cel_idx.')')
        ->setCellValue('F'.$cel_idx, $excelObj->getActiveSheet()->getCell('E'.$cel_idx)->getCalculatedValue());

    $cel_idx++;
}

//Ocultando a celula que contem a formula
$excelObj->getActiveSheet()->getColumnDimension('E')->setVisible(false);

//Congelamento de colunas - primeira linha - cabeçalho
$excelObj->getActiveSheet()->freezePane('A2');
//Congelamento de colunas - primeira linha e primeira coluna
$excelObj->getActiveSheet(0)->freezePaneByColumnAndRow(1, 2);

//Configurando o tamanho das colunas
foreach (range('A', 'F') as $columnID) {
    $excelObj->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
}

//Formatação
$excelObj->getActiveSheet()->getStyle('A1:F2560')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
$excelObj->getActiveSheet()->getStyle('A1:F2560')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setWrapText(true);

//Filtro
$excelObj->getActiveSheet()->getStyle('A1:F1')->getFont()->setBold(true);
$excelObj->getActiveSheet()->setAutoFilter($excelObj->getActiveSheet()->calculateWorksheetDimension());

//Cores - Fonte
$excelObj->getActiveSheet()->getStyle('A1:B1')->getFont()->setColor(new PHPExcel_Style_Color(PHPExcel_Style_Color::COLOR_YELLOW));
$excelObj->getActiveSheet()->getStyle('C1:E1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_GREEN);
$excelObj->getActiveSheet()->getStyle('F1')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);

//Cores - Fundo
$excelObj->getActiveSheet()->getStyle('A1:F1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB(PHPExcel_Style_Color::COLOR_BLACK);

//var_dump($excelObj);

//Escreve e salva o documento
$objWriter = PHPExcel_IOFactory::createWriter($excelObj, 'Excel2007');
$objWriter->save($filename);

echo "<br />[FIM]<br />";
echo "</pre>";

?>
