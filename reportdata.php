<?php
include('koneksi.php');
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
 
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Nama');
$sheet->setCellValue('B1', 'Jenis_Kel');
$sheet->setCellValue('C1', 'NISN');
$sheet->setCellValue('D1', 'NIK');
$sheet->setCellValue('E1', 'Kota');
$sheet->setCellValue('F1', 'Lahir');
$sheet->setCellValue('G1', 'No_Akta');
$sheet->setCellValue('H1', 'Agama');
$sheet->setCellValue('I1', 'Negara');
$sheet->setCellValue('J1', 'Ber_khusus');
$sheet->setCellValue('K1', 'Alamat');
$sheet->setCellValue('L1', 'RT');
$sheet->setCellValue('M1', 'RW');
$sheet->setCellValue('N1', 'Dusun');
$sheet->setCellValue('O1', 'Kelurahan');
$sheet->setCellValue('P1', 'Kecamatan');
$sheet->setCellValue('Q1', 'Kode_Pos');
$sheet->setCellValue('R1', 'Lintang');
$sheet->setCellValue('S1', 'Bujur');
$sheet->setCellValue('T1', 'Tempat');
$sheet->setCellValue('U1', 'Kendaraan');
$sheet->setCellValue('V1', 'Anak_ke');
$sheet->setCellValue('W1', 'Penerima_KKS');
$sheet->setCellValue('X1', 'NO_KKS');

$i=2;
$query = mysqli_query($conn,"SELECT * FROM pendaftaran");
while($row = mysqli_fetch_array($query))
{
	$sheet->setCellValue('A'.$i, $row['nama']);
	$sheet->setCellValue('B'.$i, $row['jenis_kel']);
	$sheet->setCellValue('C'.$i, $row['nisn']);
	$sheet->setCellValue('D'.$i, $row['nik']);
	$sheet->setCellValue('E'.$i, $row['kota_lahir']);	
	$sheet->setCellValue('F'.$i, $row['tanggal_lahir']);
	$sheet->setCellValue('G'.$i, $row['no_akta']);
	$sheet->setCellValue('H'.$i, $row['agama']);
	$sheet->setCellValue('I'.$i, $row['kewarganegaraan']);
	$sheet->setCellValue('J'.$i, $row['berkebutuhan_khusus']);
	$sheet->setCellValue('K'.$i, $row['alamat']);
	$sheet->setCellValue('L'.$i, $row['rt']);
	$sheet->setCellValue('M'.$i, $row['rw']);
	$sheet->setCellValue('N'.$i, $row['dusun']);
	$sheet->setCellValue('O'.$i, $row['kelurahan']);
	$sheet->setCellValue('P'.$i, $row['kecamatan']);
	$sheet->setCellValue('Q'.$i, $row['kode_pos']);
	$sheet->setCellValue('R'.$i, $row['lintang']);
	$sheet->setCellValue('S'.$i, $row['bujur']);
	$sheet->setCellValue('T'.$i, $row['tinggal']);
	$sheet->setCellValue('U'.$i, $row['kendaraan']);
	$sheet->setCellValue('V'.$i, $row['anak']);
	$sheet->setCellValue('W'.$i, $row['penerima']);
	$sheet->setCellValue('X'.$i, $row['KKS']);	
	
	$i++;				
}
 
$styleArray = [
			'borders' => [
				'allBorders' => [
					'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
				],
			],
		];
$i = $i - 1;
$sheet->getStyle('A1:X'.$i)->applyFromArray($styleArray);
 
$writer = new Xlsx($spreadsheet);
$writer->save('Report Data Siswa.xlsx');
?>