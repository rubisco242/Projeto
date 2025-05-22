<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;

define('INPUT_FILE', 'C:\Users\anton\Downloads\SL7000 761_7,26_90128151_LP1_2025_05_20_17_24_33.xlsx');
define('OUTPUT_FILE', 'ficheiro_saida.sgl_v2');

// ======================
// Funções de formatação
// ======================

function formatField(string $valor, int $comprimento, string $caracterPreenchimento = ' ', int $tipoPreenchimento = STR_PAD_RIGHT): string {
    return str_pad(substr($valor, 0, $comprimento), $comprimento, $caracterPreenchimento, $tipoPreenchimento);
}

function formatNumericField(string $valor, int $comprimento, string $caracterPreenchimento = '0'): string {
    $valor = preg_replace('/\\D/', '', $valor);
    return str_pad(substr($valor, -$comprimento), $comprimento, $caracterPreenchimento, STR_PAD_LEFT);
}

// ======================
// Funções de Registos SGL_V2
// ======================

function criarRegisto00(array $dados): string {
    return '00' .
        formatField($dados['entidade'], 10) .
        formatField($dados['data'], 8) .
        formatField($dados['sequencia'], 4) .
        str_repeat(' ', 76); // preenchimento até 100 chars
}

function criarRegisto20(array $dados): string {
    return '20' .
        formatField($dados['cpe'], 20) .
        formatField($dados['data'], 8) .
        formatField($dados['hora'], 4) .
        formatNumericField($dados['energia'], 10) .
        formatField($dados['tipoLeitura'], 1) .
        str_repeat(' ', 57); // completar até 100
}

function criarRegisto99(int $total): string {
    return '99' .
        formatField('TOTAL', 10) .
        formatNumericField((string) $total, 8) .
        str_repeat(' ', 80);
}

// ======================
// Código Principal
// ======================

try {
    $spreadsheet = IOFactory::load(INPUT_FILE);
    $sheet = $spreadsheet->getSheetByName('Load Profile data');

    $rows = $sheet->toArray();
    $dataAtual = null;
    $registos = [];

    $cpe = 'PT123456789012345678'; // substituir com o verdadeiro CPE
    $tipoLeitura = 'R';
    $entidade = 'EGAC00123';
    $sequencia = '0001';
    $dataHoje = date('Ymd');

    // Cabeçalho
    $registos[] = criarRegisto00([
        'entidade' => $entidade,
        'data' => $dataHoje,
        'sequencia' => $sequencia
    ]);

    // Ignorar cabeçalhos
    for ($i = 1; $i < count($rows); $i++) {
        $row = $rows[$i];

        // Atualiza a data se estiver na primeira coluna
        if (!empty($row[0])) {
            if (is_numeric($row[0])) {
                $excelDate = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($row[0]);
            } else {
                $excelDate = \DateTime::createFromFormat('d/m/Y', $row[0]);
            }

            if ($excelDate !== false) {
                $dataAtual = $excelDate->format('Ymd');
            }
        }

        // Ignora se ainda não tiver uma data válida
        if (!$dataAtual) {
            continue;
        }

        // Lê a hora da segunda coluna
        $horaStr = trim($row[1]);
        $horaObj = \DateTime::createFromFormat('H:i', $horaStr);
        if (!$horaObj) {
            continue; // hora inválida
        }
        $hora = $horaObj->format('Hi'); // "HHMM"

        // Energia está na terceira coluna
        $energia = (float)$row[2];
        $energiaFormatada = number_format($energia * 100, 0, '', '');

        $registos[] = criarRegisto20([
            'cpe' => $cpe,
            'data' => $dataAtual,
            'hora' => $hora,
            'energia' => $energiaFormatada,
            'tipoLeitura' => $tipoLeitura
        ]);
    }



    // Rodapé
    $totalDetalhes = count($registos) - 1; // exceto o cabeçalho
    $registos[] = criarRegisto99($totalDetalhes);

    // Escrever ficheiro
    $conteudo = implode(PHP_EOL, $registos);
    file_put_contents(OUTPUT_FILE, mb_convert_encoding($conteudo, 'ISO-8859-1'));

    echo "Ficheiro SGL_V2 gerado com sucesso!\\n";

} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
    echo "Erro ao ler o ficheiro Excel ;(: " . $e->getMessage();
} catch (Exception $e) {
    echo "Erro inesperado (linha {$e->getLine()}): " . $e->getMessage();
}
