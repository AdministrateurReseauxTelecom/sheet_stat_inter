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
	$sql= 	"SELECT llx_user.lastname, COUNT(llx_fichinter.rowid)
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


function calcul_tab_vente ($duree, $pdo)
{
	//statut 0 brouillon 1 valider 2 signé 3 perdu 4 facturé
	$tableau =array(); 
	$ligne = array("C1" => "Statut", "C2" => "Nombre", "C3" => "Montant HT");
	array_push($tableau, $ligne);		
	
	//nombre et montant de proposition par utilisateur et par statut
	
			
	$sql= 	"SELECT COUNT(p.rowid), SUM(p.total_ht), p.fk_statut 
			FROM llx_propal AS p
			WHERE 	(".$duree."(p.datec)=".$duree."(NOW()) 
					AND year(p.datec)=YEAR(NOW()) 
					AND p.fk_statut <= 1 ) 
					OR (".$duree."(p.date_cloture)=".$duree."(NOW()) 
					AND year(p.date_cloture)=YEAR(NOW()) 
					AND p.fk_statut > 1)
			GROUP BY p.fk_statut";
			
	
	foreach  ($pdo->query($sql) as $row) 
	{
		//guits debug
    	//$tt = dol_buildpath("/fichinter/card.php?id=".$object->id, 1);
    	
		//$arr = get_defined_vars(); //affiche toutes les variables
		ob_start(); 

		var_export($row); 

		$tab_debug=ob_get_contents(); 
		ob_end_clean(); 
		$fichier=fopen('tes_xls.log','w'); 
		fwrite($fichier,$tab_debug); 
		fclose($fichier); 
		//guits debug fin
		switch ($row['fk_statut'])
		{
			case '0':
				$ligne = array("C1" => "Brouilon", "C2" => $row['COUNT(p.rowid)'], "C3" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);				
			break;
			case '1':
				$ligne = array("C1" => "Validée", "C2" => $row['COUNT(p.rowid)'], "C3" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '2':
				$ligne = array("C1" => "Signée", "C2" => $row['COUNT(p.rowid)'], "C3" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '4':
				$ligne = array("C1" => "Facturée", "C2" => $row['COUNT(p.rowid)'], "C3" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '3':
				$ligne = array("C1" => "Perdu", "C2" => $row['COUNT(p.rowid)'], "C3" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
		}	
	
	}
			
	return $tableau;	
}


function calcul_tab_vente_user ($duree, $pdo)
{
	//statut 0 brouillon 1 valider 2 signé 3 perdu 4 facturé
	$tableau =array(); 
	$ligne = array("C1" => "Nom", "C2" => "Statut", "C3" => "Nombre", "C4" => "Montant HT");
	array_push($tableau, $ligne);		
	
	//nombre et montant de proposition par utilisateur et par statut
	
			
	$sql= 	"SELECT u.lastname, COUNT(p.rowid), SUM(p.total_ht), p.fk_statut 
			FROM llx_user AS u, llx_propal AS p INNER JOIN llx_element_contact AS c ON p.rowid = c.element_id 
			WHERE 	(u.rowid =c.fk_socpeople 
					AND c.fk_c_type_contact=31
					AND ".$duree."(p.datec)=".$duree."(NOW()) 
					AND year(p.datec)=YEAR(NOW()) 
					AND p.fk_statut <= 1 ) 
					OR (u.rowid =c.fk_socpeople 
					AND c.fk_c_type_contact=31
					AND ".$duree."(p.date_cloture)=".$duree."(NOW()) 
					AND year(p.date_cloture)=YEAR(NOW()) 
					AND p.fk_statut > 1)
			GROUP BY u.lastname, p.fk_statut";
			
	
	foreach  ($pdo->query($sql) as $row) 
	{
		//guits debug
    	//$tt = dol_buildpath("/fichinter/card.php?id=".$object->id, 1);
    	
		//$arr = get_defined_vars(); //affiche toutes les variables
		ob_start(); 

		var_export($row); 

		$tab_debug=ob_get_contents(); 
		ob_end_clean(); 
		$fichier=fopen('tes_xls.log','w'); 
		fwrite($fichier,$tab_debug); 
		fclose($fichier); 
		//guits debug fin
		switch ($row['fk_statut'])
		{
			case '0':
				$ligne = array("C1" => $row['lastname'], "C2" => "Brouilon", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);				
			break;
			case '1':
				$ligne = array("C1" => $row['lastname'], "C2" => "Validée", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '2':
				$ligne = array("C1" => $row['lastname'], "C2" => "Signée", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '4':
				$ligne = array("C1" => $row['lastname'], "C2" => "Facturée", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '3':
				$ligne = array("C1" => $row['lastname'], "C2" => "Perdu", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
		}	
	
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



//Fonction qui reçois un tableau de données, un sheet, un numéro de ligne, la premiere colonne qui accueillent le tableau, la largeur (en nbcolonne) de la premiere colonne, le nombre de colonnes et un titre de tableau
function print_tableau ($tableau, $sheet, $C1, $i, $largeur_1, $nbcolonne, $titre_tableau)
{	
	$ligne = array ();
	
	$colonne = ord ($C1) + $largeur_1 -1;
	$C2 = chr ($colonne);
	$C3 =  chr ($colonne + $nbcolonne -1);

	
	//entete du tableau
	$sheet->mergeCells($C1.$i.':'.$C3.$i);		//fusion de cellules
	$sheet->setCellValue($C1.$i, $titre_tableau);
	$sheet->getStyle('A1')->getAlignment()->setWrapText(true);		
	$sheet->getStyle($C1.$i.':'.$C3.$i)->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C1.$i.':'.$C3.$i)->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C1.$i)->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$sheet->getStyle($C3.$i)->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
	$i++;
	
	//corps du tableau	
	foreach ($tableau as $ligne)
	{
		if ($largeur_1 > 1)
		{
			$sheet->mergeCells($C1.$i.':'.$C2.$i);		//fusion de cellules
		}
		$sheet->setCellValue($C1.$i, $ligne['C1']);
		$numC = 2;
		$colonne = ord ($C2) + 1;
		$colonne = chr ($colonne);
		while ( $numC <= $nbcolonne)
		{
			$sheet->setCellValue($colonne.$i, $ligne['C'.$numC]);
			$sheet->getStyle($colonne.$i)->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
			$numC++;
			$colonne = ord ($colonne) + 1;
			$colonne = chr ($colonne);
		}
		
		//Bordures du tableau, (BORDER_THICK pour les traits en gras)
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







// CREATE A NEW SHEET + POPULATE DATA

$sheet = $spreadsheet->createSheet();
//$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Statistiques commerciales');

$sheet->mergeCells('A1:F1');		//fusion de cellules
$sheet->setCellValue('A1', "Statistiques commerciales du ".date('d')."/".date('m')."/".date('Y')." (semaine n°".date('W').")");


$my_tab = calcul_tab_vente ('WEEK', $pdo);
print_tableau ($my_tab, $sheet, 'A', 4, 2, 3, "Ventes de la semaine");

$my_tab = calcul_tab_vente ('MONTH', $pdo);
print_tableau ($my_tab, $sheet, 'F', 4, 2, 3, "Ventes du mois");

$my_tab = calcul_tab_vente ('YEAR', $pdo);
print_tableau ($my_tab, $sheet, 'K', 4, 2, 3, "Ventes de l'année");

$my_tab = calcul_tab_vente_user ('WEEK', $pdo);
print_tableau ($my_tab, $sheet, 'A', 13, 2, 4, "Ventes de la semaine par utilisateur");

$my_tab = calcul_tab_vente_user ('MONTH', $pdo);
print_tableau ($my_tab, $sheet, 'G', 13, 2, 4, "Ventes du mois par utilisateur");

$my_tab = calcul_tab_vente_user ('YEAR', $pdo);
print_tableau ($my_tab, $sheet, 'M', 13, 2, 4, "Ventes de l'année par utilisateur");
//$colonne = ord ('A');
//$colonne++;
//$colonne = chr ($colonne);
//$sheet->setCellValue('A3', $colonne);
//$test_tab = array ();
//$t = 1;

//$test_tab_ligne = array ( 'C1'=>"C1_".$t, 'C2'=>"C2_".$t, 'C3'=>"C3_".$t, 'C4'=>"C4_".$t);
//array_push($test_tab, $test_tab_ligne);
//$t++;
//$test_tab_ligne = array ( 'C1'=>"C1_".$t, 'C2'=>"C2_".$t, 'C3'=>"C3_".$t, 'C4'=>"C4_".$t);
//array_push($test_tab, $test_tab_ligne);
//$t++;
//$test_tab_ligne = array ( 'C1'=>"C1_".$t, 'C2'=>"C2_".$t, 'C3'=>"C3_".$t, 'C4'=>"C4_".$t);
//array_push($test_tab, $test_tab_ligne);
//$t++;
//$test_tab_ligne = array ( 'C1'=>"C1_".$t, 'C2'=>"C2_".$t, 'C3'=>"C3_".$t, 'C4'=>"C4_".$t);
//array_push($test_tab, $test_tab_ligne);
//$t++;
//$test_tab_ligne = array ( 'C1'=>"C1_".$t, 'C2'=>"C2_".$t, 'C3'=>"C3_".$t, 'C4'=>"C4_".$t);
//array_push($test_tab, $test_tab_ligne);
//$t++;

//print_tableau ($test_tab, $sheet, 'B', 8, 3, 4, "Interventions validées cette année\npar utilisateur");


//SELECT u.lastname, 
//SUM( IF( MONTH( p.date_cloture ) =1, total_ht, NULL ) ) AS janv, 
//SUM( IF( MONTH( p.date_cloture ) =2, total_ht, NULL ) ) AS fevr, 
//SUM( IF( MONTH( p.date_cloture ) =3, total_ht, NULL ) ) AS mars, 
//SUM( IF( MONTH( p.date_cloture ) =4, total_ht, NULL ) ) AS avril, 
//SUM( IF( MONTH( p.date_cloture ) =5, total_ht, NULL ) ) AS mai, 
//SUM( IF( MONTH( p.date_cloture ) =6, total_ht, NULL ) ) AS juin, 
//SUM( IF( MONTH( p.date_cloture ) =7, total_ht, NULL ) ) AS juill, 
//SUM( IF( MONTH( p.date_cloture ) =8, total_ht, NULL ) ) AS aout, 
//SUM( IF( MONTH( p.date_cloture ) =9, total_ht, NULL ) ) AS sept, 
//SUM( IF( MONTH( p.date_cloture ) =10, total_ht, NULL ) ) AS oct, 
//SUM( IF( MONTH( p.date_cloture ) =11, total_ht, NULL ) ) AS nov, 
//SUM( IF( MONTH( p.date_cloture ) =12, total_ht, NULL ) ) AS dece 
//FROM llx_user AS u, llx_propal AS p INNER JOIN llx_element_contact AS c ON p.rowid = c.element_id 
//WHERE u.rowid =c.fk_socpeople AND p.fk_statut =2 AND year(p.date_cloture)=2017 AND c.fk_c_type_contact=31 
//GROUP BY c.fk_socpeople










// OUTPUT vesrion fichier sur disque dur
$spreadsheet->getProperties()
    ->setTitle('Statistiques interventions')
    ->setSubject('Statistiques interventions')
    ->setDescription('Statistiques interventions par semeine et par mois')
    ->setCreator('A.R.T.')
    ->setLastModifiedBy('A.R.T.');
$writer = new Xlsx($spreadsheet);
//$writer->save("/var/www/html/dolibarrdelta/documents/ecm/rapports/Stats_inter_". gmdate('D, d M Y H:i:s').".xlsx");
//$writer->save("Stats_inter_". gmdate('D, d M Y H:i:s').".xlsx");
$writer->save(gmdate('Ymd')."_Stats_inter_.xlsx");
//$writer->save("/var/www/html/dolibarrdelta/documents/ecm/rapports/".gmdate('Ymd')."_Stats_hebdo_.xlsx");













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

