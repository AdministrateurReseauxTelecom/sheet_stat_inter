<?php
// CONNECT TO DATABASE
define('DB_HOST', 'localhost');
define('DB_NAME', 'dolibarrdelta');
define('DB_CHARSET', 'utf8');
define('DB_USER', 'root');
define('DB_PASSWORD', 'mysql');
$pdo = new PDO(
  "mysql:host=".DB_HOST.";dbname=".DB_NAME.";charset=".DB_CHARSET, 
  DB_USER, DB_PASSWORD, [
    PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
    PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
    PDO::ATTR_EMULATE_PREPARES => false,
  ]
);


// CREATE PHPSPREADSHEET OBJECT
require "vendor/autoload.php";
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;



//SQL

//$sql = "SELECT COUNT(llx_fichinter.rowid), REPLACE ( REPLACE (llx_fichinter.fk_statut, '3', 'Traitées'), '1', 'Validées') FROM llx_fichinter WHERE llx_fichinter.date_valid >= (NOW() - INTERVAL 1 WEEK) group by llx_fichinter.fk_statut";

//~ $sql = "SELECT COUNT(llx_fichinter.rowid),";
//~ $sql."REPLACE ( REPLACE (llx_fichinter.fk_statut, '3', 'Traitées'), '1', 'Validées')";
//~ $sql."FROM llx_fichinter WHERE llx_fichinter.date_valid >= (NOW() - INTERVAL 1000 WEEK) group by llx_fichinter.fk_statut";








// CREATE A NEW SPREADSHEET + POPULATE DATA
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Statistiques interventions');



$sql = "SELECT COUNT(llx_fichinter.rowid), (llx_fichinter.fk_statut) FROM llx_fichinter WHERE ((llx_fichinter.date_valid >= (NOW() - INTERVAL 1 WEEK)) AND (llx_fichinter.fk_statut >= 1)) OR (llx_fichinter.fk_statut = 0) group by llx_fichinter.fk_statut";


$i = 2;
foreach  ($pdo->query($sql) as $row) 
{
	switch ($row['fk_statut'])
	{
		case '0':
			$sheet->setCellValue('B'.$i, $row['COUNT(llx_fichinter.rowid)']);
			$sheet->setCellValue('A'.$i, 'Brouillon');
			
			$sheet->getStyle('A'.$i)->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
			$sheet->getStyle('A'.$i)->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
			$sheet->getStyle('A'.$i)->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
			$sheet->getStyle('A'.$i)->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
			
			$sheet->getStyle('B'.$i)->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
			$sheet->getStyle('B'.$i)->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
			$sheet->getStyle('B'.$i)->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
			$sheet->getStyle('B'.$i)->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK);
						
			
		break;
		case '1':
			$sheet->setCellValue('B'.$i, $row['COUNT(llx_fichinter.rowid)']);
			$sheet->setCellValue('A'.$i, 'Validées');
		break;
	}
	
	
	
        //$sheet->setCellValue('A'.$i, $row['COUNT(llx_fichinter.rowid)']);
		//$sheet->setCellValue('B'.$i, $row['fk_statut']);
		$i++;
    
  }

//$sql = "SELECT COUNT(llx_fichinter.rowid), (llx_fichinter.fk_statut) FROM llx_fichinter WHERE ((llx_fichinter.date_valid >= (NOW() - INTERVAL 1 month)) AND (llx_fichinter.fk_statut >= 1)) OR (llx_fichinter.fk_statut = 0) group by llx_fichinter.fk_statut";
$sql=	"SELECT COUNT(llx_fichinter.rowid), (llx_fichinter.fk_statut) 
		FROM llx_fichinter 
		WHERE ((MONTH(llx_fichinter.date_valid) >= MONTH(NOW()))
			   AND (YEAR(llx_fichinter.date_valid) >= YEAR(NOW()))
			   AND (llx_fichinter.fk_statut >= 1)) 
			   OR (llx_fichinter.fk_statut = 0) 
			   group by llx_fichinter.fk_statut";

$i = 10;
foreach  ($pdo->query($sql) as $row) {
	//guits debug
    	
		//$arr = get_defined_vars(); //affiche toutes les variables
		ob_start(); 

		var_export($row); 

		$tab_debug=ob_get_contents(); 
		ob_end_clean(); 
		$fichier=fopen('tes_xls.log','w'); 
		fwrite($fichier,$tab_debug); 
		fclose($fichier); 
		//guits debug fin
        $sheet->setCellValue('A'.$i, $row['COUNT(llx_fichinter.rowid)']);
		$sheet->setCellValue('B'.$i, $row['fk_statut']);
		$i++;
    
  }


// OUTPUT vesrion fichier sur disque dur
$spreadsheet->getProperties()
    ->setTitle('Statistiques interventions')
    ->setSubject('Statistiques interventions')
    ->setDescription('Statistiques interventions par semeine et par mois')
    ->setCreator('A.R.T.')
    ->setLastModifiedBy('A.R.T.');
$writer = new Xlsx($spreadsheet);
$writer->save("/var/www/html/dolibarrdelta/documents/ecm/rapports/Stats_inter_". gmdate('D, d M Y H:i:s').".xlsx");





// OUTPUT vesrion fichier à enregistre via navigateur
//$writer = new Xlsx($spreadsheet);
//header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//header('Content-Disposition: attachment;filename="users.xlsx"');
//header('Cache-Control: max-age=0');
//header('Expires: Fri, 11 Nov 2011 11:11:11 GMT');
//header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
//header('Cache-Control: cache, must-revalidate');
//header('Pragma: public');
//$writer->save('php://output');



?>
