<?php

error_reporting(E_ALL);
ini_set('display_errors', true);
ini_set('display_startup_errors', true);

class EasyExcel
{

    private $excelObj;
    private $sheet;
    private $fname;
    private $writeObj;
    private $dataSample;
    private $operationsSample;

    public function __construct($filename, $phpexcel)
    {
        if(!$phpexcel) {
            die("\n<br />Error, PHPExcel path not defined: This param is required to create instance of object<br />\n");
        }

        if(!$filename) {
            die("\n<br />Error, filename not defined: This param is required to file create<br />\n");
        }

        require_once($phpexcel);

        $this->excelObj = new PHPExcel();
        $this->sheet = $this->excelObj->setActiveSheetIndex(0);
        $this->fname = $filename;

        return $this;
    }

    public function getDataSample()
    {

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

        $this->dataSample = $results;

        return $this->dataSample;
    }

    public function getOperationsSample()
    {

        // SAMPLE:
        // F2=E2=SUM(C2:D2)
        // F2=E2=AVERAGE(C2:D2)

        $operations = array(
            0 => ['oper' => '=SUM', 'start' => 'B', 'end' => 'D', 'ftarget' => 'E', 'rtarget' => 'F', 'mode' => 'repeat', 'char' => ':'],
            1 => ['oper' => '=AVERAGE', 'start' => 'B', 'end' => 'D', 'ftarget' => 'G', 'rtarget' => 'H', 'mode' => 'repeat', 'char' => ':']
        );

        $this->operationsSample = $operations;

        return $this->operationsSample;

    }

    public function setFileProps($params = array([0 => '', 1 => '', 2 => '', 3 => '', 4 => '', 5 => '', 6 => '']))
    {

        $this->excelObj->getProperties()
            ->setCreator($params[0])
            ->setLastModifiedBy($params[1])
            ->setTitle($params[2])
            ->setSubject($params[3])
            ->setDescription($params[4])
            ->setKeywords($params[5])
            ->setCategory($params[6]);

        return $this->excelObj;

    }

    public function setXlsHeader($params = array())
    {

        $idx = 0;

        foreach (range('A', 'Z') as $col) {

            if (isset($params[$idx])) {

                $this->sheet->setCellValue($col . '1', $params[$idx]);
            }

            $idx++;

        }

        return $this->sheet;

    }

    public function setXlsContent($idx, $params = array(), $operations = array())
    {

        //Verifica se foram requisitadas operacoes matematicas
        if (count($operations) > 0 && (!$this->excelObj || $this->excelObj == false)) {

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

                    $this->sheet->setCellValue($l . $idx, $v);

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
                        $this->sheet->setCellValue($fl . $idx, "=SUM(" . $ri . $idx . $ch . $rf . $idx . ")");
                        $this->sheet->setCellValue($rt . $idx, $this->excelObj->getActiveSheet()->getCell($fl . $idx)->getCalculatedValue());
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
                        $this->sheet->setCellValue($fl . $idx, "=AVERAGE(" . $ri . $idx . $ch . $rf . $idx . ")");
                        $this->sheet->setCellValue($rt . $idx, $this->excelObj->getActiveSheet()->getCell($fl . $idx)->getCalculatedValue());
                    }
                }
            }

            $idx++;
        }

        /*echo "<h3>RESPONSE</h3>";
        echo $s;
        echo "<hr />";*/

        return $this->sheet;

    }

    public function setXlsStyles($params = array())
    {

        /*Ocultar colunas*/

        if (isset($params['hidden_column']) && $params['hidden_column'] != "") {

            $cols = explode(",", $params['hidden_column']);
            if (count($cols) > 0) {
                for ($y = 0; $y < count($cols); $y++) {
                    $this->excelObj->getActiveSheet()->getColumnDimension($cols[$y])->setVisible(false);
                }
            } else {
                $this->excelObj->getActiveSheet()->getColumnDimension($params['hidden_column'])->setVisible(false);
            }

        }

        /*Congelar linhas e colunas*/

        //Congelamento de colunas - primeira linha - cabeçalho
        if (isset($params['freeze_row']) && $params['freeze_row'] != "") {
            $this->excelObj->getActiveSheet()->freezePane($params['freeze_row']);
        }

        //Congelamento de colunas - primeira linha e primeira coluna
        if (isset($params['freeze_byrow']) && $params['freeze_byrow'] != "" && $params['freeze_bycolumn'] != "") {
            $this->excelObj->getActiveSheet(0)->freezePaneByColumnAndRow($params['freeze_bycolumn'], $params['freeze_byrow']);
        }

        /*Largura de colunas: AutoSize*/

        //Configurando o tamanho de todas as colunas com autosize
        if (isset($params['autosize_all']) && $params['autosize_all'] == '1') {

            foreach (range('A', 'Z') as $columnID) {
                $this->excelObj->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
            }

        }

        //Configurando o tamanho de apenas uma coluna com autosize
        if (isset($params['autosize_uniq']) && $params['autosize_uniq'] != "") {

            $this->excelObj->getActiveSheet()->getColumnDimension($params['autosize_uniq'])->setAutoSize(true);

        }

        //Configurando o tamanho de um range decolunas com autosize
        if (isset($params['autosize_range_ini']) && $params['autosize_range_ini'] != "" && $params['autosize_range_fin'] != "") {

            foreach (range($params['autosize_range_ini'], $params['autosize_range_fin']) as $columnID) {
                $this->excelObj->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
            }
        }

        /*Formatação*/

        //Alinhamento Horizontal
        if (isset($params['horizontal_align']) && $params['horizontal_align'] != "" && $params['cells_align'] != "") {
            if ($params['horizontal_align'] == "center") {
                $this->excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)->setWrapText(true);
            }
            if ($params['horizontal_align'] == "left") {
                $this->excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setWrapText(true);
            }
            if ($params['horizontal_align'] == "right") {
                $this->excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT)->setWrapText(true);
            }
            if ($params['horizontal_align'] == "justify") {
                $this->excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY)->setWrapText(true);
            }
        }

        //Alinhamento Vertical
        if (isset($params['vertical_align']) && $params['vertical_align'] != "" && $params['cells_align'] != "") {
            if ($params['vertical_align'] == "center") {
                $this->excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)->setWrapText(true);
            }
            if ($params['vertical_align'] == "left") {
                $this->excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT)->setWrapText(true);
            }
            if ($params['vertical_align'] == "right") {
                $this->excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT)->setWrapText(true);
            }
            if ($params['vertical_align'] == "justify") {
                $this->excelObj->getActiveSheet()->getStyle($params['cells_align'])->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_JUSTIFY)->setWrapText(true);
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

                $this->excelObj->getActiveSheet()->getStyle($range)->getFont()->getColor()->setARGB($color);

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

                $this->excelObj->getActiveSheet()->getStyle($range)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB($color);

            }
        }

        /*Filtro*/
        if (isset($params['filter']) && $params['filter'] != "") {
            $this->excelObj->getActiveSheet()->getStyle($params['filter'])->getFont()->setBold(true);
            $this->excelObj->getActiveSheet()->setAutoFilter($this->excelObj->getActiveSheet()->calculateWorksheetDimension());
        }

        return $this->excelObj;

    }

    public function saveAs($opt)
    {

        $this->writeObj = PHPExcel_IOFactory::createWriter($this->excelObj, 'Excel2007');
        $this->writeObj->save($this->fname);

        if ($opt == "excelObj") {
            return $this->excelObj;
        }
        return $this->writeObj;
    }
}

?>

