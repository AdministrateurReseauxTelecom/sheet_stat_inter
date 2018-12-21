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




//Fonction qui reçois une durée sous forme de string du et récupère les données pour retourner un tableau. variable $duree (MONTH, WEEK)
function calcul_tab ($duree, $pdo)
{
	$tableau =array();
	
	//Nombre d'intervention enregistrée
	$sql=	"SELECT COUNT(llx_fichinter.rowid) 
			FROM llx_fichinter 
			WHERE (((".$duree." (llx_fichinter.datec)) = ( ".$duree." ( NOW())))
			AND ((YEAR (llx_fichinter.datec)) = ( YEAR ( NOW()))))";
			
	foreach  ($pdo->query($sql) as $row) 
	{
		$res = $row['COUNT(llx_fichinter.rowid)'];
	}
	
	$ligne = array("Titre" => "Enregistrée", "Nb" => $res);
	array_push($tableau, $ligne);
	
	
	//Nombre d'intervention par statut
	$sql=	"SELECT COUNT(llx_fichinter.rowid), (llx_fichinter.fk_statut) 
			FROM llx_fichinter 
			WHERE ((".$duree."(llx_fichinter.date_valid) >= ".$duree."(NOW()))
					AND (YEAR(llx_fichinter.date_valid) >= YEAR(NOW()))
					AND (llx_fichinter.fk_statut >= 1)) 
				OR ((llx_fichinter.fk_statut = 0)
					AND (((".$duree." (llx_fichinter.datec)) = ( ".$duree." ( NOW())))
					AND ((YEAR (llx_fichinter.datec)) = ( YEAR ( NOW())))))
				group by llx_fichinter.fk_statut";
				
	foreach  ($pdo->query($sql) as $row) 
	{
		switch ($row['fk_statut'])
		{
			case '0':
				$ligne = array("Titre" => "Brouillon", "Nb" => $row['COUNT(llx_fichinter.rowid)']);
				array_push($tableau, $ligne);				
			break;
			case '1':
				$ligne = array("Titre" => "Validée", "Nb" => $row['COUNT(llx_fichinter.rowid)']);
				array_push($tableau, $ligne);
			break;
			case '3':
				$ligne = array("Titre" => "Cloturée", "Nb" => $row['COUNT(llx_fichinter.rowid)']);
				array_push($tableau, $ligne);
			break;
			case '5':
				$ligne = array("Titre" => "Facturée", "Nb" => $row['COUNT(llx_fichinter.rowid)']);
				array_push($tableau, $ligne);
			break;
		}			
	  }

		return $tableau;
}

//Fonction qui reçois une durée sous forme de string du et récupère les données pour retourner un tableau. variable $duree (MONTH, WEEK)
function calcul_tab_user ($duree, $pdo)
{
	$tableau =array();
	
	//Nombre d'intervention Validé par technicien
	$sql= 	"SELECT COUNT(llx_fichinter.rowid), llx_user.lastname
			FROM `llx_fichinter` 
			LEFT JOIN llx_user ON llx_fichinter.fk_user_valid=llx_user.rowid
			WHERE ((".$duree."(llx_fichinter.date_valid) >= ".$duree."(NOW()))
			AND (YEAR(llx_fichinter.date_valid) >= YEAR(NOW()))
			AND (llx_fichinter.fk_statut >= 1)) 
			GROUP BY llx_user.lastname";
	
	foreach  ($pdo->query($sql) as $row) 
	{
		$ligne = array("Titre" => $row['lastname'], "Nb" => $row['COUNT(llx_fichinter.rowid)']);
		array_push($tableau, $ligne);				
	}

		return $tableau;	
}

//Fonction qui reçois un tableau de données, un sheet, un numéro de ligne et les 3 colonnes qui accueillent le tableau + un numéro de ligne
function print_tab ($tableau, $sheet, $C1, $C2, $C3, $i, $durée, $titre_tableau)
{	
	$ligne = array ();
	
	//entete du premeir tableau
	$sheet->mergeCells($C1.$i.':'.$C3.$i);		//fusion de cellules
	$sheet->setCellValue($C1.$i, $titre_tableau);
	$sheet->getStyle('A1')->getAlignment()->setWrapText(true);		
	$sheet->getStyle($C1.$i.':'.$C3.$i)->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C1.$i.':'.$C3.$i)->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C1.$i)->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C3.$i)->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$i++;
	$sheet->mergeCells($C1.$i.':'.$C2.$i);		//fusion de cellules
	$sheet->setCellValue($C1.$i, "Statut");	
	$sheet->getStyle($C1.$i)->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C1.$i)->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C1.$i)->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C1.$i)->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->setCellValue($C3.$i, "Nombre");	
	$sheet->getStyle($C3.$i)->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C3.$i)->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C3.$i)->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C3.$i)->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	
	
	foreach ($tableau as $ligne)
	{
		$sheet->mergeCells($C1.$i.':'.$C2.$i);		//fusion de cellules
		$sheet->setCellValue($C1.$i, $ligne['Titre']);
		$sheet->setCellValue($C3.$i, $ligne['Nb']);
		
		//BORDER_THIN pour les traits en gras
		$sheet->getStyle($C1.$i.':'.$C3.$i)->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
		$sheet->getStyle($C1.$i.':'.$C3.$i)->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
		$sheet->getStyle($C1.$i)->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
		$sheet->getStyle($C2.$i)->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
		$sheet->getStyle($C3.$i)->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
		$i++;
	}
}


		
// CREATE A NEW SPREADSHEET + POPULATE DATA
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Statistiques interventions');

$sheet->mergeCells('A1:F1');		//fusion de cellules
$sheet->setCellValue('A1', "Statistiques interventions du ".date('d')."/".date('m')."/".date('Y')." (semaine n°".date('W').")");

//Nombre d'intervention enregistrée
$sql=	"SELECT COUNT(llx_fichinter.rowid) 
		FROM llx_fichinter 
		WHERE llx_fichinter.fk_statut = 0";
		
foreach  ($pdo->query($sql) as $row) 
{
	$res = $row['COUNT(llx_fichinter.rowid)'];
}
$sheet->mergeCells('A3:F3');		//fusion de cellules
$sheet->setCellValue('A3', "Nombre total d'interventions brouillon : ".$res);


$my_tab = calcul_tab ('WEEK', $pdo);
print_tab ($my_tab, $sheet, 'A', 'B', 'C', 5, 'WEEK', "Interventions de la semaine");

$my_tab = calcul_tab ('MONTH', $pdo);
print_tab ($my_tab, $sheet, 'E', 'F', 'G', 5, 'MONTH', "Interventions du mois");

$my_tab = calcul_tab ('YEAR', $pdo);
print_tab ($my_tab, $sheet, 'I', 'J', 'K', 5, 'YEAR', "Interventions de l'année");

$sheet->getRowDimension('13')->setRowHeight(30);
$my_tab = calcul_tab_user ('WEEK', $pdo);
print_tab ($my_tab, $sheet, 'A', 'C', 'D', 13, 'WEEK', "Interventions validées cette semaine\npar utilisateur");

$my_tab = calcul_tab_user ('MONTH', $pdo);
print_tab ($my_tab, $sheet, 'F', 'H', 'I', 13, 'MONTH', "Interventions validées ce mois\npar utilisateur");

$my_tab = calcul_tab_user ('YEAR', $pdo);
print_tab ($my_tab, $sheet, 'K', 'M', 'N', 13, 'YEAR', "Interventions validées cette année\npar utilisateur");





// OUTPUT vesrion fichier sur disque dur
$spreadsheet->getProperties()
    ->setTitle('Statistiques interventions')
    ->setSubject('Statistiques interventions')
    ->setDescription('Statistiques interventions par semeine et par mois')
    ->setCreator('A.R.T.')
    ->setLastModifiedBy('A.R.T.');
$writer = new Xlsx($spreadsheet);
//$writer->save("/var/www/html/dolibarrdelta/documents/ecm/rapports/Stats_inter_". gmdate('D, d M Y H:i:s').".xlsx");
$writer->save("Stats_inter_". gmdate('D, d M Y H:i:s').".xlsx");








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
