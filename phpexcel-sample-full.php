<?php

error_reporting(E_ALL);
ini_set('display_errors', true);
ini_set('display_startup_errors', true);

//Importa a lib para trabalhar com PHPExcel
include("./lib/sistema/PHPExcel/Classes/PHPExcel/IOFactory.php");

function getData() {

    $results = array(
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

    return $results;
}

function getOperations() {

    // SAMPLE:
    // F2=E2=SUM(C2:D2)
    // F2=E2=AVERAGE(C2:D2)

    $operations = array(
        0 => ['oper' => '=SUM', 'start' => 'B', 'end' => 'D', 'ftarget' => 'E', 'rtarget' => 'F', 'mode' => 'repeat', 'char' => ':'],
        1 => ['oper' => '=AVERAGE', 'start' => 'B', 'end' => 'D', 'ftarget' => 'G', 'rtarget' => 'H', 'mode' => 'repeat', 'char' => ':']
    );

    return $operations;

}

function setFileProps($excelObj, $params = array([0 => '', 1 => '', 2 => '', 3 => '', 4 => '', 5 => '', 6 => ''])) {

    $excelObj->getProperties()
        ->setCreator($params[0])
        ->setLastModifiedBy($params[1])
        ->setTitle($params[2])
        ->setSubject($params[3])
        ->setDescription($params[4])
        ->setKeywords($params[5])
        ->setCategory($params[6]);

    return $excelObj;

}

function setXlsHeader($sheet, $params = array()) {

    $idx = 0;

    foreach (range('A', 'Z') as $col) {

        if (isset($params[$idx])) {

            $sheet->setCellValue($col . '1', $params[$idx]);
        }

        $idx++;

    }

    return $sheet;

}

function setXlsContent($sheet, $idx, $params = array(), $operations = array(), $excelObj) {

    //Verifica se foram requisitadas operacoes matematicas
    if (count($operations) > 0 && (!$excelObj || $excelObj == false)) {

        //Para operacoes é obrigatorio um objeto instanciado do PHPExcel
        die("\n<br />Error, object not defined: This param is required when operations are used<br />\n");

    }

    $s = "";

    //Lendo o resultado por linhas
    for ($n = 0; $n < count($params); $n++) {

        $s .= "<br />[{$n}] ";

        //Coluna inicial
        $l = 'A';

        //Lendo o resultado por colunas
        foreach ($params[$n] as $k => $v) {

            if (isset($k)) {

                $s .= $l . $idx . " = " . $v . "; ";

                $sheet->setCellValue($l . $idx, $v);

                //Atualizando as colunas por letras
                //de acordo com as colunas do resultado
                foreach (range($l, $l) as $c) {
                    $l++;
                }
            }
        }

        for ($y = 0; $y < count($operations); $y++) {

            //SUM (SOMA) (+)
            if ($operations[$y]['oper'] == "=SUM") {

                $ri = $operations[$y]['start'];
                $rf = $operations[$y]['end'];
                $fl = $operations[$y]['ftarget'];
                $rt = $operations[$y]['rtarget'];
                $md = $operations[$y]['mode'];
                $ch = $operations[$y]['char'];

                if ($md == "repeat") {
                    $s .= "=SUM(" . $ri . $idx . $ch . $rf . $idx . "); ";
                    $sheet->setCellValue($fl . $idx, "=SUM(" . $ri . $idx . $ch . $rf . $idx . ")");
                    $sheet->setCellValue($rt . $idx, $excelObj->getActiveSheet()->getCell($fl . $idx)->getCalculatedValue());
                }
            }

            //AVERAGE (MEDIA)
            if ($operations[$y]['oper'] == "=AVERAGE") {

                $ri = $operations[$y]['start'];
                $rf = $operations[$y]['end'];
                $fl = $operations[$y]['ftarget'];
                $rt = $operations[$y]['rtarget'];
                $md = $operations[$y]['mode'];
                $ch = $operations[$y]['char'];

                if ($md == "repeat") {
                    $s .= "=AVERAGE(" . $ri . $idx . $ch . $rf . $idx . "); ";
                    $sheet->setCellValue($fl . $idx, "=AVERAGE(" . $ri . $idx . $ch . $rf . $idx . ")");
                    $sheet->setCellValue($rt . $idx, $excelObj->getActiveSheet()->getCell($fl . $idx)->getCalculatedValue());
                }
            }
        }

        $idx++;
    }

    /*echo "<h3>RESPONSE</h3>";
    echo $s;
    echo "<hr />";*/

    return $sheet;

}

function setXlsStyles($excelObj, $params = array()) {

    /*Ocultar colunas*/

    if (isset($params['hidden_column']) && $params['hidden_column'] != "") {

        $cols = explode(",", $params['hidden_column']);
        if (count($cols) > 0) {
            for ($y = 0; $y < count($cols); $y++) {
                $excelObj->getActiveSheet()->getColumnDimension($cols[$y])->setVisible(false);
            }
        } else {
            $excelObj->getActiveSheet()->getColumnDimension($params['hidden_column'])->setVisible(false);
        }

    }

    /*Congelar linhas e colunas*/

    //Congelamento de colunas - primeira linha - cabeçalho
    if (isset($params['freeze_row']) && $params['freeze_row'] != "") {
        $excelObj->getActiveSheet()->freezePane($params['freeze_row']);
    }

    //Congelamento de colunas - primeira linha e primeira coluna
    if (isset($params['freeze_byrow']) && $params['freeze_byrow'] != "" && $params['freeze_bycolumn'] != "") {
        $excelObj->getActiveSheet(0)->freezePaneByColumnAndRow($params['freeze_bycolumn'], $params['freeze_byrow']);
    }

    /*Largura de colunas: AutoSize*/

    //Configurando o tamanho de todas as colunas com autosize
    if (isset($params['autosize_all']) && $params['autosize_all'] == '1') {

        foreach (range('A', 'Z') as $columnID) {
            $excelObj->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
        }

    }

    //Configurando o tamanho de apenas uma coluna com autosize
    if (isset($params['autosize_uniq']) && $params['autosize_uniq'] != "") {

        $excelObj->getActiveSheet()->getColumnDimension($params['autosize_uniq'])->setAutoSize(true);

    }

    //Configurando o tamanho de um range decolunas com autosize
    if (isset($params['autosize_range_ini']) && $params['autosize_range_ini'] != "" && $params['autosize_range_fin'] != "") {

        foreach (range($params['autosize_range_ini'], $params['autosize_range_fin']) as $columnID) {
            $excelObj->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
        }
    }

    /*Formatação*/

    //Alinhamento Horizontal
    if (isset($params['horizontal_align']) && $params['horizontal_align'] != "" && $params['cells_align'] != "") {
        if ($params['horizontal_align'] == "center") {
            $excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
        }
        if ($params['horizontal_align'] == "left") {
            $excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setWrapText(true);
        }
        if ($params['horizontal_align'] == "right") {
            $excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT)->setWrapText(true);
        }
        if ($params['horizontal_align'] == "justify") {
            $excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY)->setWrapText(true);
        }
    }

    //Alinhamento Vertical
    if (isset($params['vertical_align']) && $params['vertical_align'] != "" && $params['cells_align'] != "") {
        if ($params['vertical_align'] == "center") {
            $excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setWrapText(true);
        }
        if ($params['vertical_align'] == "left") {
            $excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setWrapText(true);
        }
        if ($params['vertical_align'] == "right") {
            $excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT)->setWrapText(true);
        }
        if ($params['vertical_align'] == "justify") {
            $excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY)->setWrapText(true);
        }
    }

    /*Cores*/

    //Fonte
    if (isset($params['fonte_color']) && is_array($params['fonte_color']) && count($params['fonte_color']) > 0) {
        $color = "";
        for ($i = 0; $i < count($params['fonte_color']); $i++) {

            $range = $params['fonte_color'][$i]['cells_range'];
            $color = $params['fonte_color'][$i]['cells_color'];

            if ($color == 'yellow') {
                $color = PHPExcel_Style_Color::COLOR_YELLOW;
            } elseif ($color == 'red') {
                $color = PHPExcel_Style_Color::COLOR_RED;
            } elseif ($color == 'green') {
                $color = PHPExcel_Style_Color::COLOR_GREEN;
            } elseif ($color == 'black') {
                $color = PHPExcel_Style_Color::COLOR_BLACK;
            } elseif ($color == 'blue') {
                $color = PHPExcel_Style_Color::COLOR_BLUE;
            }

            $excelObj->getActiveSheet()->getStyle($range)->getFont()->getColor()->setARGB($color);

        }
    }

    //Fundo
    if (isset($params['fill_color']) && is_array($params['fill_color']) && count($params['fill_color']) > 0) {
        $color = "";
        for ($i = 0; $i < count($params['fill_color']); $i++) {

            $range = $params['fill_color'][$i]['cells_range'];
            $color = $params['fill_color'][$i]['cells_color'];

            if ($color == 'yellow') {
                $color = PHPExcel_Style_Color::COLOR_YELLOW;
            } elseif ($color == 'red') {
                $color = PHPExcel_Style_Color::COLOR_RED;
            } elseif ($color == 'green') {
                $color = PHPExcel_Style_Color::COLOR_GREEN;
            } elseif ($color == 'black') {
                $color = PHPExcel_Style_Color::COLOR_BLACK;
            } elseif ($color == 'blue') {
                $color = PHPExcel_Style_Color::COLOR_BLUE;
            }

            $excelObj->getActiveSheet()->getStyle($range)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB($color);

        }
    }

    /*Filtro*/
    if (isset($params['filter']) && $params['filter'] != "") {
        $excelObj->getActiveSheet()->getStyle($params['filter'])->getFont()->setBold(true);
        $excelObj->getActiveSheet()->setAutoFilter($excelObj->getActiveSheet()->calculateWorksheetDimension());
    }

    return $excelObj;

}

function saveAs($fname, $excelObj, $opt) {

    $writerObj = PHPExcel_IOFactory::createWriter($excelObj, 'Excel2007');
    $writerObj->save($fname);

    if ($opt == "excelObj") {
        return $excelObj;
    }
    return $writerObj;
}

echo "<pre>";
echo "<br />[INICIO]<br />";

//Estilos e Formatações da planilha
$styles   = array(
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

//Propiedades do documento
$props  = array('Jereelton','Jereelton','Teste','Teste Subject','Teste Description','Teste Keywords','Teste Category');

//Header da planilha
$header = array('Nome Campo 1','Nome Campo 2','Nome Campo 3','Nome Campo 4','Formula','Total','Formula','Media');

//Dados da planilha
$param    = getData();echo "<br />[param]<br />";
$oper     = getOperations();echo "<br />[oper]<br />";
$fname    = "Arquivo_Teste_".date('dmYHis').".xlsx";echo "<br />[fname]<br />";

//Instancia do objeto PHPExcel
$excelObj = new PHPExcel();echo "<br />[PHPExcel]<br />";
$excelObj = setFileProps($excelObj, $props);echo "<br />[setFileProps]<br />";

//Escrita dos dados: Cabeçalho e Conteudo
$idx   = 2;
$sheet = $excelObj->setActiveSheetIndex(0);echo "<br />[setActiveSheetIndex]<br />";
$sheet = setXlsHeader($sheet, $header);echo "<br />[setXlsHeader]<br />";
$sheet = setXlsContent($sheet, $idx, $param, $oper, $excelObj);echo "<br />[setXlsContent]<br />";

//Formatando a planilha: Opcional
$excelObj = setXlsStyles($excelObj, $styles);echo "<br />[setXlsStyles]<br />";

//Grava e Salva o arquivo XLSX
$excelObj = saveAs($fname, $excelObj, 'excelObj');echo "<br />[saveAs]<br />";

echo "Arquivo: ". $fname;
echo "<br />[FIM]<br />";
echo "</pre>";

?>

