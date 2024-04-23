<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

$reader = new Xlsx();
$spreadsheet = $reader->load("PLANO DE CONTAS.xlsx");

$conn = new mysqli('localhost', 'root', '', 'jvaz');
if ($conn->connect_error) {
    die("Falha na conexão: " . $conn->connect_error);
}

$worksheet = $spreadsheet->getActiveSheet();

foreach ($worksheet->getRowIterator() as $row) {
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(false);
    $cells = [];

    foreach ($cellIterator as $cell) {
        $cells[] = $cell->getValue();
    }

    if ($row->getRowIndex() > 1) {
        $conta = $conn->real_escape_string($cells[0]);
        $tipoConta = $conn->real_escape_string($cells[1]);
        $nome = $conn->real_escape_string($cells[2]);
        $natureza = $conn->real_escape_string($cells[3]);

        $stmt = $conn->prepare("INSERT INTO tb_contas (id, conta, tipo_conta, nome, natureza) VALUES (NULL, ?, ?, ?, ?)");
        $stmt->bind_param('isss', $conta, $tipoConta, $nome, $natureza);
        $stmt->execute();
    }
}

$conn->close();
?>