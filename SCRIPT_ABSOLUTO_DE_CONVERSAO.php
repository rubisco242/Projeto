<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;

define('INPUT_FILE', 'C:\Users\anton\Downloads\SL7000 761_7,26_90128151_LP1_2025_05_20_17_24_33.xlsx');

// === Função para gerar nome de ficheiro ===
function gerarNomeFicheiro($cpe, $sequencial = 1) {
    $prefixo = 'ACBTM0';
    $aleatorio = 'CEL' . substr(str_shuffle('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'), 0, 3);
    $data = date('Ymd');
    return "{$prefixo}{$aleatorio}PE{$cpe}_{$data}_{$sequencial}.sgl_v2";
}

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

    $entidade = 'EDIS';
    $perfil = 'POTENCIA';
    $intervalo = '15M';
    $cpe = 'PT0002970085950235BP';
    $sequencial = 1;
    $dataHoje = date('Ymd');
    $dataAtual = null;

    // === Nomes genéricos dos canais (6 canais) ===
    $nomesCanais = ['A+', 'Ri+', 'Rc-'];
    // Remove nomes vazios ou espaços
    $nomesCanais = array_filter($nomesCanais, fn($nome) => trim($nome) !== '');

    // === Número total de canais ===
    $numCanais = count($nomesCanais);

    // === REGISTO 20 ===
    $registos20 = [];
    $datasExcel = [];

    for ($i = 1; $i < count($rows); $i++) {
        $row = $rows[$i];

        if (!empty($row[0])) {
            $dataObj = is_numeric($row[0]) ? Date::excelToDateTimeObject($row[0]) : \DateTime::createFromFormat('d/m/Y', $row[0]);
            if ($dataObj) {
                $dataAtual = $dataObj->format('Ymd');
                $datasExcel[] = $dataAtual;
            }
        }

        if (!$dataAtual || empty($row[1])) {
            continue;
        }

        $horaObj = \DateTime::createFromFormat('H:i', trim($row[1]));
        if (!$horaObj) continue;

        $hora = $horaObj->format('Hi');
        $linha = ['20', $dataAtual, $hora];

        for ($c = 2; $c < 2 + $numCanais; $c++) {
            $valor = isset($row[$c]) ? $row[$c] : 0;
            $linha[] = formatEnergia($valor);
            $linha[] = '0';
            $linha[] = '000000000000.000';
        }

        $registos20[] = implode("\t", $linha);
    }

    // Datas início e fim do bloco 00
    $dataInicio = min($datasExcel);
    $dataFim = max($datasExcel);

    // === REGISTO 00 ===
    $registos[] = implode("\t", [
        '00',
        $entidade,
        '0000/0',
        '0000000001',
        '0000000000',
        '00000001',
        $dataInicio,
        $dataFim,
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

    // === REGISTO 04 === (nomes canais todos) ===
    $registos[] = implode("\t", array_merge(['04'], array_slice($nomesCanais, 0, $numCanais)));

    // === Adiciona os registos 20 ===
    $registos = array_merge($registos, $registos20);

    // === REGISTO 99 + prefixo 000000 sem "TOTAL" e tudo na mesma linha ===
    $totalLinhas = count($registos) + 1;
    $registos[] = implode("\t", ['99', '000000', str_pad($numCanais, 6, '0', STR_PAD_LEFT), str_pad($totalLinhas, 6, '0', STR_PAD_LEFT)]);

    $outputFile = gerarNomeFicheiro($cpe, $sequencial);
    $conteudo = implode(PHP_EOL, $registos) . PHP_EOL;
    file_put_contents($outputFile, mb_convert_encoding($conteudo, 'ISO-8859-1'));

    echo "Ficheiro SGL_V2 gerado com sucesso: $outputFile" . PHP_EOL;

} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
    echo "Erro ao ler o ficheiro Excel: " . $e->getMessage();
} catch (Exception $e) {
    echo "Erro inesperado na linha {$e->getLine()}: " . $e->getMessage();
}



