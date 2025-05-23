<?php
 
require 'vendor/autoload.php';
 
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
 
define('INPUT_FILE', 'C:\Users\anton\Downloads\SL7000 761_7,26_90128151_LP1_2025_05_20_17_24_33.xlsx');  // <-- Ajusta o caminho do ficheiro Excel aqui
define('OUTPUT_FILE', 'ficheiro_saida.sgl_v2');
 
// === Formatação da energia ===
function formatEnergia($valor) {
    $valorMil = round((float)$valor * 1000);
    return '000000000000.' . str_pad($valorMil, 3, '0', STR_PAD_LEFT);
}
 
try {
    $spreadsheet = IOFactory::load(INPUT_FILE);
    $sheet = $spreadsheet->getSheetByName('Load Profile data');
    $rows = $sheet->toArray();
 
    $registos = [];
 
    // === CONFIGURAÇÕES (Não sei qual é a entidade certa) ===
    $dataHoje = date('Ymd');
    $entidade = 'EDIS';
    $perfil = 'POTENCIA';
    $intervalo = '15M'; // Intervalo de 15 minutos
    $numCanais = 6;
    $dataAtual = null;
 
    // === REGISTO 00 (Com a primeira e a última data do Excel)=== 
    $registos[] = implode("\t", [
        '00',
        $entidade,
        '0000/0',
        '0000000001',
        '0000000000',
        '00000001',
        '20250401',
        '20250430',
        '002',
        '999999',
        '00',
        '00',
        $dataHoje
    ]);
 
    // === REGISTO 01 ===
    $registos[] = implode("\t", [
        '01', 'D', 'S', '01', $perfil, 'K', $intervalo, '1', '1', 'P'
    ]);
 
    // === REGISTO 04 ===
    $registos[] = implode("\t", ['04', 'A+', 'Ri+', 'Rc-', 'A-', 'Ri-', 'Rc+']);  // Neste caso para 6 canais (acho eu)
 
    // === REGISTOS 20 ===
    for ($i = 1; $i < count($rows); $i++) {
        $row = $rows[$i];
 
        // Datas (coluna A)
        if (!empty($row[0])) {
            $dataObj = is_numeric($row[0]) ? Date::excelToDateTimeObject($row[0]) : \DateTime::createFromFormat('d/m/Y', $row[0]);
            if ($dataObj) {
                $dataAtual = $dataObj->format('Ymd');
            }
        }
 
        if (!$dataAtual || empty($row[1])) {
            continue;
        }
        // Horas em intervalos de 15 minutos (coluna B)
        $horaObj = \DateTime::createFromFormat('H:i', trim($row[1]));
        if (!$horaObj) continue;
 
        $hora = $horaObj->format('Hi');
 
        // Montar as linhas do tipo 20
        $linha = ['20', $dataAtual, $hora];
 
        for ($c = 2; $c <= 7; $c++) {
            $valor = isset($row[$c]) ? $row[$c] : 0;
            $linha[] = formatEnergia($valor);
            $linha[] = '0';  // Flag sempre 0
        }
 
        $registos[] = implode("\t", $linha);
    }
 
    // === REGISTO 99 (retira-se 1 para não contar o cabecalho)===
    $registos[] = implode("\t", ['99', 'TOTAL', count($registos) - 1]);
 
    // === GRAVAR FICHEIRO ===
    $conteudo = implode(PHP_EOL, $registos);
    file_put_contents(OUTPUT_FILE, mb_convert_encoding($conteudo, 'ISO-8859-1'));
 
    echo "✅ Ficheiro SGL_V2 gerado com sucesso: " . OUTPUT_FILE . PHP_EOL;
 
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
    echo "Erro ao ler o ficheiro Excel: " . $e->getMessage();
} catch (Exception $e) {
    echo "Erro inesperado na linha {$e->getLine()}: " . $e->getMessage();
}