<?php
// CONNECT TO DATABASE
define('DB_HOST', 'localhost');
define('DB_NAME', 'dolibarrdelta');
//define('DB_NAME', 'dolitest');
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


//Fonction qui liste les devis perdu
//reçois la durée sous forme de string (MONTH, WEEK, YEAR)
//recois un pdo
function liste_devis_gagne ($duree, $pdo)
{
	$tableau =array();
	$ligne = array("C1" => "N° devis", "C2" => "Rèf. client", "C3" => "Client", "C4" => "Auteur", "C5" => "Date de création", "C6" => "Date de cloture", "C7" => "Montant HT", "C8" => "Type de prestation");
	array_push($tableau, $ligne);
	
	//liste de devis perdus triée par datec et user, le user est l'auteur ou le commercial resp suivi
	$sql="SELECT llx_propal.ref,
				llx_propal.ref_client,
				llx_societe.nom,
				llx_user.firstname,
				llx_user.lastname, 
				llx_propal.datec,
				llx_propal.date_cloture,
				llx_propal.total_ht,
				llx_propal_extrafields.ztypepresta
			FROM llx_propal 
				LEFT JOIN llx_propal_extrafields ON (llx_propal_extrafields.fk_object=llx_propal.rowid)
				LEFT JOIN llx_societe ON (llx_societe.rowid=llx_propal.fk_soc)
				LEFT JOIN llx_element_contact ON (llx_element_contact.element_id=llx_propal.rowid)
				LEFT JOIN llx_user ON (llx_user.rowid=llx_element_contact.fk_socpeople)
				WHERE ((llx_element_contact.fk_c_type_contact=31) 
						AND llx_propal.fk_statut=2
						AND YEAR(llx_propal.date_cloture)>=YEAR(NOW())
						AND ".$duree."(llx_propal.date_cloture)>=".$duree."(NOW()))
						OR
						((llx_element_contact.fk_c_type_contact=31) 
						AND llx_propal.fk_statut=4
						AND YEAR(llx_propal.date_cloture)>=YEAR(NOW())
						AND ".$duree."(llx_propal.date_cloture)>=".$duree."(NOW()))
			UNION
			SELECT llx_propal.ref,
				llx_propal.ref_client,
				llx_societe.nom,
				llx_user.firstname,
				llx_user.lastname, 
				llx_propal.datec,
				llx_propal.date_cloture,
				llx_propal.total_ht,
			llx_propal_extrafields.ztypepresta
			FROM llx_propal 
					LEFT JOIN llx_propal_extrafields ON (llx_propal_extrafields.fk_object=llx_propal.rowid) 
					LEFT JOIN llx_user ON (llx_user.rowid=llx_propal.fk_user_author) 
					LEFT JOIN llx_societe ON (llx_societe.rowid=llx_propal.fk_soc)
					WHERE (llx_propal.rowid NOT IN (SELECT llx_element_contact.element_id FROM llx_element_contact WHERE llx_element_contact.fk_c_type_contact=31)
						AND llx_propal.fk_statut=2
						AND YEAR(llx_propal.date_cloture)=YEAR(NOW())
						AND ".$duree."(llx_propal.date_cloture)=".$duree."(NOW()))
						OR
						(llx_propal.rowid NOT IN (SELECT llx_element_contact.element_id FROM llx_element_contact WHERE llx_element_contact.fk_c_type_contact=31)
						AND llx_propal.fk_statut=4
						AND YEAR(llx_propal.date_cloture)=YEAR(NOW())
						AND ".$duree."(llx_propal.date_cloture)=".$duree."(NOW()))
			ORDER BY datec desc, lastname asc";
			
	
	$num_type = array (15,14,13,12,11,10,9,8,7,6,5,4,3,2,1);
	$verb_type = array ('AFF.Dynamique ','IPTV ','Wifi ','Videoprotection ', 'VDI ', 'TV ', 'Telephonie ',' Interphonie',' GTC', ' Controle d acces', ' Cablage', ' Alarme SSI', ' Alarme intrusion', ' Tertiaire' , ' Habitat');
			
	foreach  ($pdo->query($sql) as $row) 
	{
		$type = str_replace ($num_type, $verb_type, $row['ztypepresta']);
		
		$ligne = array("C1" => $row['ref'], "C2" => $row['ref_client'], "C3" => $row['nom'], "C4" => $row['firstname']." ".$row['lastname'], "C5" => $row['datec'],"C6" => $row['date_cloture'], "C7" => $row['total_ht'], "C8" => $type);
		
		array_push($tableau, $ligne);
	}
	
	return $tableau;		
}




//Fonction qui liste les devis perdu
//reçois la durée sous forme de string (MONTH, WEEK, YEAR)
//recois un pdo
function liste_devis_perdu ($duree, $pdo)
{
	$tableau =array();
	$ligne = array("C1" => "N° devis", "C2" => "Rèf. client", "C3" => "Client", "C4" => "Auteur", "C5" => "Date de création", "C6" => "Date de cloture", "C7" => "Montant HT", "C8" => "Type de prestation");
	array_push($tableau, $ligne);
	
	//liste de devis perdus triée par datec et user, le user est l'auteur ou le commercial resp suivi
	$sql="SELECT llx_propal.ref,
				llx_propal.ref_client,
				llx_societe.nom,
				llx_user.firstname,
				llx_user.lastname, 
				llx_propal.datec,
				llx_propal.date_cloture,
				llx_propal.total_ht,
				llx_propal_extrafields.ztypepresta
			FROM llx_propal 
				LEFT JOIN llx_propal_extrafields ON (llx_propal_extrafields.fk_object=llx_propal.rowid)
				LEFT JOIN llx_societe ON (llx_societe.rowid=llx_propal.fk_soc)
				LEFT JOIN llx_element_contact ON (llx_element_contact.element_id=llx_propal.rowid)
				LEFT JOIN llx_user ON (llx_user.rowid=llx_element_contact.fk_socpeople)
				WHERE (llx_element_contact.fk_c_type_contact=31) 
						AND llx_propal.fk_statut=3
						AND YEAR(llx_propal.date_cloture)>=YEAR(NOW())
						AND ".$duree."(llx_propal.date_cloture)>=".$duree."(NOW())
			UNION
			SELECT llx_propal.ref,
				llx_propal.ref_client,
				llx_societe.nom,
				llx_user.firstname,
				llx_user.lastname, 
				llx_propal.datec,
				llx_propal.date_cloture,
				llx_propal.total_ht,
			llx_propal_extrafields.ztypepresta
			FROM llx_propal 
					LEFT JOIN llx_propal_extrafields ON (llx_propal_extrafields.fk_object=llx_propal.rowid) 
					LEFT JOIN llx_user ON (llx_user.rowid=llx_propal.fk_user_author) 
					LEFT JOIN llx_societe ON (llx_societe.rowid=llx_propal.fk_soc)
					WHERE llx_propal.rowid NOT IN (SELECT llx_element_contact.element_id FROM llx_element_contact WHERE llx_element_contact.fk_c_type_contact=31)
						AND llx_propal.fk_statut=3
						AND YEAR(llx_propal.date_cloture)=YEAR(NOW())
						AND ".$duree."(llx_propal.date_cloture)=".$duree."(NOW())
			ORDER BY datec desc, lastname asc";
			
	
	$num_type = array (15,14,13,12,11,10,9,8,7,6,5,4,3,2,1);
	$verb_type = array ('AFF.Dynamique ','IPTV ','Wifi ','Videoprotection ', 'VDI ', 'TV ', 'Telephonie ',' Interphonie',' GTC', ' Controle d acces', ' Cablage', ' Alarme SSI', ' Alarme intrusion', ' Tertiaire' , ' Habitat');
			
	foreach  ($pdo->query($sql) as $row) 
	{
		//guits debug
    	//$tt = dol_buildpath("/fichinter/card.php?id=".$object->id, 1);
    	
		$arr = get_defined_vars(); //affiche toutes les variables
		ob_start(); 
	

		var_export($row); 

		$tab_debug=ob_get_contents(); 
		ob_end_clean(); 
		$fichier=fopen('tes_xls.log','w'); 
		fwrite($fichier,$tab_debug); 
		fclose($fichier); 
		//guits debug fin
		$type = str_replace ($num_type, $verb_type, $row['ztypepresta']);
		
		$ligne = array("C1" => $row['ref'], "C2" => $row['ref_client'], "C3" => $row['nom'], "C4" => $row['firstname']." ".$row['lastname'], "C5" => $row['datec'],"C6" => $row['date_cloture'], "C7" => $row['total_ht'], "C8" => $type);
		
		array_push($tableau, $ligne);
	}
	
	return $tableau;		
}

//Fonction qui liste les devis ouverts
//recois un pdo
function liste_devis_ouvert ($pdo)
{
	$tableau =array();
	$ligne = array("C1" => "N° devis", "C2" => "Rèf. client", "C3" => "Client", "C4" => "Auteur", "C5" => "Date de création", "C6" => "Montant HT", "C7" => "Type de prestation");
	array_push($tableau, $ligne);
	
	//liste de devis par statut triée par datec et user, le user est l'auteur ou le commercial resp suivi
	$sql="SELECT llx_propal.ref,
				llx_propal.ref_client,
				llx_societe.nom,
				llx_user.firstname,
				llx_user.lastname, 
				llx_propal.datec,
				llx_propal.total_ht,
				llx_propal_extrafields.ztypepresta
			FROM llx_propal 
				LEFT JOIN llx_propal_extrafields ON (llx_propal_extrafields.fk_object=llx_propal.rowid)
				LEFT JOIN llx_societe ON (llx_societe.rowid=llx_propal.fk_soc)
				LEFT JOIN llx_element_contact ON (llx_element_contact.element_id=llx_propal.rowid)
				LEFT JOIN llx_user ON (llx_user.rowid=llx_element_contact.fk_socpeople)
				WHERE (llx_element_contact.fk_c_type_contact=31) AND llx_propal.fk_statut=1 
			UNION
			SELECT llx_propal.ref,
				llx_propal.ref_client,
				llx_societe.nom,
				llx_user.firstname,
				llx_user.lastname, 
				llx_propal.datec,
				llx_propal.total_ht,
			llx_propal_extrafields.ztypepresta
			FROM llx_propal 
					LEFT JOIN llx_propal_extrafields ON (llx_propal_extrafields.fk_object=llx_propal.rowid) 
					LEFT JOIN llx_user ON (llx_user.rowid=llx_propal.fk_user_author) 
					LEFT JOIN llx_societe ON (llx_societe.rowid=llx_propal.fk_soc)
					WHERE llx_propal.rowid NOT IN (SELECT llx_element_contact.element_id FROM llx_element_contact WHERE llx_element_contact.fk_c_type_contact=31) AND  llx_propal.fk_statut=1
			ORDER BY datec desc, lastname asc";
			
		
			//REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(llx_propal_extrafields.ztypepresta,15, 'AFF.Dynamique '),14, 'IPTV '),13, 'Wifi '),12, 'Videoprotection '),11, 'VDI '),10, 'TV '),9, 'Telephonie '),8, 'Interphonie '),7, 'GTC '),6, 'Controle d acces '),5, 'Cablage '),4, 'Alarme SSI '),3, 'Alarme intrusion '),2, 'Tertiaire '),1 , 'Habitat ')
			
	//$tt = $pdo->fetchAll();
	$num_type = array (15,14,13,12,11,10,9,8,7,6,5,4,3,2,1);
	$verb_type = array ('AFF.Dynamique ','IPTV ','Wifi ','Videoprotection ', 'VDI ', 'TV ', 'Telephonie ',' Interphonie',' GTC', ' Controle d acces', ' Cablage', ' Alarme SSI', ' Alarme intrusion', ' Tertiaire' , ' Habitat');
			
	foreach  ($pdo->query($sql) as $row) 
	{
		
		
		$type = str_replace ($num_type, $verb_type, $row['ztypepresta']);
		
		//$ligne = array("C1" => $row['llx_propal.ref'], "C2" => $row['llx_propal.ref_client'], "C3" => $row['llx_societe.nom'], "C4" => $row['llx_user.firstname']." ".$row['llx_user.lastname'], "C5" => $row['llx_propal.datec'], "C6" => $row['llx_propal.total_ht'], "C7" => $row["REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(llx_propal_extrafields.ztypepresta,15, 'AFF.Dynamique '),14, 'IPTV '),13, 'Wifi '),12, 'Vidéoprotection '),11, 'VDI '),10, 'TV '),9, 'Téléphonie '),8, 'Interphonie '),7, 'GTC '),6, 'Contrôle d'acces '),5, 'Cablage '),4, 'Alarme SSI '),3, 'Alarme intrusion '),2, 'Tertiaire '),1 , 'Habitat ')"]);
		$ligne = array("C1" => $row['ref'], "C2" => $row['ref_client'], "C3" => $row['nom'], "C4" => $row['firstname']." ".$row['lastname'], "C5" => $row['datec'], "C6" => $row['total_ht'], "C7" => $type);
		
		array_push($tableau, $ligne);
	}


		



	
		return $tableau;
}


//Fonction qui compte le nombre d'interventions par statut sur une durée sélectionnée
//reçois la durée sous forme de string (MONTH, WEEK, YEAR)
//recois un pdo
function calcul_tab_inter ($duree, $pdo)
{
	$tableau =array();
	$ligne = array("C1" => "Statut", "C2" => "Nombre");
	array_push($tableau, $ligne);
	
	//Nombre d'intervention enregistrée
	$sql=	"SELECT COUNT(llx_fichinter.rowid) 
			FROM llx_fichinter 
			WHERE (((".$duree." (llx_fichinter.datec)) = ( ".$duree." ( NOW())))
			AND ((YEAR (llx_fichinter.datec)) = ( YEAR ( NOW()))))";
			
	foreach  ($pdo->query($sql) as $row) 
	{
		$res = $row['COUNT(llx_fichinter.rowid)'];
	}
	
	$ligne = array("C1" => "Enregistrée", "C2" => $res);
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
				$ligne = array("C1" => "Brouillon", "C2" => $row['COUNT(llx_fichinter.rowid)']);
				array_push($tableau, $ligne);				
			break;
			case '1':
				$ligne = array("C1" => "Validée", "C2" => $row['COUNT(llx_fichinter.rowid)']);
				array_push($tableau, $ligne);
			break;
			case '3':
				$ligne = array("C1" => "Clôturée", "C2" => $row['COUNT(llx_fichinter.rowid)']);
				array_push($tableau, $ligne);
			break;
			case '5':
				$ligne = array("C1" => "Facturée", "C2" => $row['COUNT(llx_fichinter.rowid)']);
				array_push($tableau, $ligne);
			break;
		}			
	  }

		return $tableau;
}

//Fonction qui compte le nombre d'interventions validées sur une durée sélectionnée par utilisateur
//reçois la durée sous forme de string (MONTH, WEEK, YEAR)
//recois un pdo
function calcul_tab_inter_user ($duree, $pdo)
{
	$tableau =array();
	$ligne = array("C1" => "Nom", "C2" => "Nombre");
	array_push($tableau, $ligne);	
	
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
		$ligne = array("C1" => $row['lastname'], "C2" => $row['COUNT(llx_fichinter.rowid)']);
		array_push($tableau, $ligne);				
	}

		return $tableau;	
}

//Fonction qui compte les vente sur une durée sélectionnée
//reçois la durée sous forme de string (MONTH, WEEK, YEAR)
//recois un pdo
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
		
		switch ($row['fk_statut'])
		{
			case '0':
				$ligne = array("C1" => "Brouillon", "C2" => $row['COUNT(p.rowid)'], "C3" => $row['SUM(p.total_ht)']);
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
				$ligne = array("C1" => "Perdue", "C2" => $row['COUNT(p.rowid)'], "C3" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
		}	
	
	}
			
	return $tableau;	
}

//Fonction qui compte les vente par utilisateur sur une durée sélectionnée
//reçois la durée sous forme de string (MONTH, WEEK, YEAR)
//recois un pdo
function calcul_tab_vente_user ($duree, $pdo)
{
	//statut 0 brouillon 1 valider 2 signé 3 perdu 4 facturé
	$tableau =array(); 
	$ligne = array("C1" => "Nom", "C2" => "Statut", "C3" => "Nombre", "C4" => "Montant HT");
	array_push($tableau, $ligne);		
	
	//nombre et montant de proposition par utilisateur et par statut
	
			
	$sql= 	"SELECT T.lastname, COUNT(T.rowid), SUM(T.total_ht), T.fk_statut 
				FROM
					(
						
				 (SELECT DISTINCT llx_user.lastname, llx_propal.rowid, llx_propal.total_ht, llx_propal.fk_statut 
				 FROM llx_propal 
						LEFT JOIN llx_propal_extrafields ON (llx_propal_extrafields.fk_object=llx_propal.rowid)
						LEFT JOIN llx_societe ON (llx_societe.rowid=llx_propal.fk_soc)
						LEFT JOIN llx_element_contact ON (llx_element_contact.element_id=llx_propal.rowid)
						LEFT JOIN llx_user ON (llx_user.rowid=llx_element_contact.fk_socpeople)
						WHERE (".$duree."(llx_propal.datec)=".$duree."(NOW())
								AND llx_element_contact.fk_c_type_contact=31 
								AND year(llx_propal.datec)=YEAR(NOW()) 
								AND llx_propal.fk_statut <= 1 ) 
								OR
								(".$duree."(llx_propal.date_cloture)=".$duree."(NOW())
								AND llx_element_contact.fk_c_type_contact=31
								AND year(llx_propal.date_cloture)=YEAR(NOW()) 
								AND llx_propal.fk_statut > 1))
								
				UNION

				(SELECT DISTINCT llx_user.lastname, llx_propal.rowid, llx_propal.total_ht, llx_propal.fk_statut 
								FROM llx_propal LEFT JOIN llx_user ON (llx_user.rowid=llx_propal.fk_user_author )
								WHERE 	(".$duree."(llx_propal.datec)=".$duree."(NOW())
									AND year(llx_propal.datec)=YEAR(NOW()) 
									AND llx_propal.fk_statut <= 1 ) 
									OR (".$duree."(llx_propal.date_cloture)=".$duree."(NOW())
									AND year(llx_propal.date_cloture)=YEAR(NOW()) 
									AND llx_propal.fk_statut > 1) )
					)AS T
				GROUP BY T.lastname, T.fk_statut";
			
	
	foreach  ($pdo->query($sql) as $row) 
	{
		switch ($row['fk_statut'])
		{
			case '0':
				$ligne = array("C1" => $row['lastname'], "C2" => "Brouillon", "C3" => $row['COUNT(T.rowid)'], "C4" => $row['SUM(T.total_ht)']);
				array_push($tableau, $ligne);				
			break;
			case '1':
				$ligne = array("C1" => $row['lastname'], "C2" => "Validée", "C3" => $row['COUNT(T.rowid)'], "C4" => $row['SUM(T.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '2':
				$ligne = array("C1" => $row['lastname'], "C2" => "Signée", "C3" => $row['COUNT(T.rowid)'], "C4" => $row['SUM(T.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '4':
				$ligne = array("C1" => $row['lastname'], "C2" => "Facturée", "C3" => $row['COUNT(T.rowid)'], "C4" => $row['SUM(T.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '3':
				$ligne = array("C1" => $row['lastname'], "C2" => "Perdue", "C3" => $row['COUNT(T.rowid)'], "C4" => $row['SUM(T.total_ht)']);
				array_push($tableau, $ligne);	
			break;
		}	
	
	}
			
	return $tableau;	
}


//Fonction qui compte les vente par Agence sur une durée sélectionnée
//reçois la durée sous forme de string (MONTH, WEEK, YEAR)
//recois un pdo
function calcul_tab_vente_agence ($duree, $pdo)
{
	//statut 0 brouillon 1 valider 2 signé 3 perdu 4 facturé
	$tableau =array(); 
	$ligne = array("C1" => "Agence", "C2" => "Statut", "C3" => "Nombre", "C4" => "Montant HT");
	array_push($tableau, $ligne);		
	
	//nombre et montant de proposition par catégorie et par statut
	//requete a revoir!!!!
			
	$sql= 	"SELECT x.zagence, COUNT(p.rowid), SUM(p.total_ht), p.fk_statut 
			FROM llx_propal_extrafields AS x, llx_propal AS p
			WHERE 	(x.fk_object =p.rowid
					AND ".$duree."(p.date_cloture)=".$duree."(NOW()) 
					AND year(p.date_cloture)=YEAR(NOW()) 
					AND p.fk_statut > 1)
					OR
					(x.fk_object =p.rowid
					AND ".$duree."(p.datec)=".$duree."(NOW()) 
					AND year(p.datec)=YEAR(NOW()) 
					AND p.fk_statut <= 1 )
			GROUP BY x.zagence, p.fk_statut";
			
	
	foreach  ($pdo->query($sql) as $row) 
	{
		switch ($row['zagence'])
		{
			case '1':
				$agence = "Marseille";
			break;
			case '2':
				$agence = "Le Thor";	
			break;
		}
				
		switch ($row['fk_statut'])
		{
			case '0':
				$ligne = array("C1" => $agence, "C2" => "Brouillon", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);				
			break;
			case '1':
				$ligne = array("C1" => $agence, "C2" => "Validée", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '2':
				$ligne = array("C1" => $agence, "C2" => "Signée", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '4':
				$ligne = array("C1" => $agence, "C2" => "Facturée", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '3':
				$ligne = array("C1" => $agence, "C2" => "Perdue", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
		}	
	
	}
			
	return $tableau;	
}

//Fonction qui compte les vente par categorie sur une durée sélectionnée
//reçois la durée sous forme de string (MONTH, WEEK, YEAR)
//recois un pdo
function calcul_tab_vente_categ ($duree, $pdo)
{
	//statut 0 brouillon 1 valider 2 signé 3 perdu 4 facturé
	$tableau =array(); 
	$ligne = array("C1" => "Type", "C2" => "Statut", "C3" => "Nombre", "C4" => "Montant HT");
	array_push($tableau, $ligne);		
	
	//nombre et montant de proposition par catégorie et par statut
	//requete a revoir!!!!
			
	$sql= 	"SELECT x.ztypepresta, COUNT(p.rowid), SUM(p.total_ht), p.fk_statut 
			FROM llx_propal_extrafields AS x, llx_propal AS p
			WHERE 	(x.fk_object =p.rowid
					AND ".$duree."(p.date_cloture)=".$duree."(NOW()) 
					AND year(p.date_cloture)=YEAR(NOW()) 
					AND p.fk_statut > 1)
					OR
					(x.fk_object =p.rowid
					AND ".$duree."(p.datec)=".$duree."(NOW()) 
					AND year(p.datec)=YEAR(NOW()) 
					AND p.fk_statut <= 1 )
			GROUP BY x.ztypepresta, p.fk_statut";
			
			
			
	
	foreach  ($pdo->query($sql) as $row) 
	{
		
		

		
		switch ($row['ztypepresta'])
		{
			case '1':
				$typepresta = "Habitat";
			break;
			case '2':
				$typepresta = "Tertiaire";	
			break;
			case '3':
				$typepresta = "Alarme intrusion";	
			break;
			case '4':
				$typepresta = "Alarme SSI";	
			break;
			case '5':
				$typepresta = "Cablage";	
			break;
			case '6':
				$typepresta = "Controle d'acces";	
			break;
			case '7':
				$typepresta = "GTC";	
			break;
			case '8':
				$typepresta = "Interphonie";	
			break;
			case '9':
				$typepresta = "Telephonie";	
			break;
			case '10':
				$typepresta = "TV";	
			break;
			case '11':
				$typepresta = "VDI";	
			break;
			case '12':
				$typepresta = "Vidéoprotection";	
			break;
			case '13':
				$typepresta = "Wifi";	
			break;
		}
		
		switch ($row['fk_statut'])
		{
			case '0':
				$ligne = array("C1" => $typepresta, "C2" => "Brouillon", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);				
			break;
			case '1':
				$ligne = array("C1" => $typepresta, "C2" => "Validée", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '2':
				$ligne = array("C1" => $typepresta, "C2" => "Signée", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '4':
				$ligne = array("C1" => $typepresta, "C2" => "Facturée", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
			case '3':
				$ligne = array("C1" => $typepresta, "C2" => "Perdue", "C3" => $row['COUNT(p.rowid)'], "C4" => $row['SUM(p.total_ht)']);
				array_push($tableau, $ligne);	
			break;
		}	
	
	}
			
	return $tableau;	
}




//Fonction qui affiche un tableau dans une feuille Excel
//reçois un tableau de données, $tableau array [C1=>a, C2=>b, C3...)
//un sheet, 
//un numéro de ligne, $i int
//la lettre de la premiere colonne qui accueillent le tableau
//la largeur (en nbcolonne) de la premiere colonne $largeur_1 int
//e nombre de colonnes $nbcolonne int
//un titre de tableau $titre_tableau string
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



//Crétation du fichier et feuille commerciale		
// CREATE A NEW SPREADSHEET + POPULATE DATA
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setTitle('Liste devis validés');

//$sheet->mergeCells('A1:F1');		//fusion de cellules
//$sheet->setCellValue('A1', "Statistiques interventions du ".date('d')."/".date('m')."/".date('Y')." (semaine n°".date('W').")");

////Nombre d'intervention enregistrée
//$sql=	"SELECT COUNT(llx_fichinter.rowid) 
		//FROM llx_fichinter 
		//WHERE llx_fichinter.fk_statut = 0";
		
//foreach  ($pdo->query($sql) as $row) 
//{
	//$res = $row['COUNT(llx_fichinter.rowid)'];
//}
$sheet->mergeCells('A1:F1');		//fusion de cellules
$sheet->setCellValue('A1', "Liste des devis validés : ");

$largeur_titre = 2;
$nb_colonne = 7;
$i = 3;
$my_tab = liste_devis_ouvert ($pdo);
print_tableau ($my_tab, $sheet, 'A', $i, $largeur_titre, $nb_colonne, "Liste des devis validés");

////Crétation du fichier et feuille commerciale perdu		
//// CREATE A NEW SPREADSHEET + POPULATE DATA
$sheet = $spreadsheet->createSheet();
$sheet->setTitle('Liste devis perdus');

$sheet->mergeCells('A1:F1');		//fusion de cellules
$sheet->setCellValue('A1', "Liste des devis perdus : ");

$largeur_titre = 2;
$nb_colonne = 8;
$i = 3;
$my_tab = liste_devis_perdu ('YEAR', $pdo);
print_tableau ($my_tab, $sheet, 'A', $i, $largeur_titre, $nb_colonne, "Liste des devis perdus");

////Crétation du fichier et feuille commerciale gagnes		
//// CREATE A NEW SPREADSHEET + POPULATE DATA
$sheet = $spreadsheet->createSheet();
$sheet->setTitle('Liste devis gagnés');

//->getColumnDimension('A')->setAutoSize(true);
$sheet->getColumnDimension('A')->setWidth(100);		//largeur de colonne
//$sheet->mergeCells('A1:F1');		//fusion de cellules
//$sheet->setCellValue('A1', "Liste des devis gagnés : ");

$largeur_titre = 2;
$nb_colonne = 8;
$i = 1;
$my_tab = liste_devis_gagne ('YEAR', $pdo);
print_tableau ($my_tab, $sheet, 'A', $i, $largeur_titre, $nb_colonne, "Liste des devis gagnés");

//$my_tab = calcul_tab_inter ('MONTH', $pdo);
//print_tableau ($my_tab, $sheet, 'E', $i, $largeur_titre, $nb_colonne, "Interventions du mois");

//$my_tab = calcul_tab_inter ('YEAR', $pdo);
//print_tableau ($my_tab, $sheet, 'I', $i, $largeur_titre, $nb_colonne, "Interventions de l'année");

//$i = 13;
//$sheet->getRowDimension($i)->setRowHeight(30);
//$my_tab = calcul_tab_inter_user ('WEEK', $pdo);
//print_tableau ($my_tab, $sheet, 'A', $i, $largeur_titre, $nb_colonne, "Interventions validées cette semaine\npar utilisateur");

//$my_tab = calcul_tab_inter_user ('MONTH', $pdo);
//print_tableau ($my_tab, $sheet, 'F', $i, $largeur_titre, $nb_colonne, "Interventions validées ce mois\npar utilisateur");

//$my_tab = calcul_tab_inter_user ('YEAR', $pdo);
//print_tableau ($my_tab, $sheet, 'K', $i, $largeur_titre, $nb_colonne, "Interventions validées cette année\npar utilisateur");






//CRéation de la feuille commerce
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




//$sheet = $spreadsheet->createSheet();
////$sheet = $spreadsheet->getActiveSheet();
//$sheet->setTitle('Statistiques commerciales suite');

//$sheet->mergeCells('A1:F1');		//fusion de cellules
//$sheet->setCellValue('A1', "Statistiques commerciales suite du ".date('d')."/".date('m')."/".date('Y')." (semaine n°".date('W').")");


//$my_tab = calcul_tab_vente_agence ('WEEK', $pdo);
//print_tableau ($my_tab, $sheet, 'A', 4, 2, 3, "Ventes par de la semaine par agence");

//$my_tab = calcul_tab_vente_agence ('MONTH', $pdo);
//print_tableau ($my_tab, $sheet, 'F', 4, 2, 3, "Ventes du mois par agence");

//$my_tab = calcul_tab_vente_agence ('YEAR', $pdo);
//print_tableau ($my_tab, $sheet, 'K', 4, 2, 3, "Ventes de l'année par agence");

//$my_tab = calcul_tab_vente_categ ('WEEK', $pdo);
//print_tableau ($my_tab, $sheet, 'A', 13, 2, 4, "Ventes de la semaine par type de prestation");

//$my_tab = calcul_tab_vente_categ ('MONTH', $pdo);
//print_tableau ($my_tab, $sheet, 'G', 13, 2, 4, "Ventes du mois par type de prestation");

//$my_tab = calcul_tab_vente_categ ('YEAR', $pdo);
//print_tableau ($my_tab, $sheet, 'M', 13, 2, 4, "Ventes de l'année par type de prestation");





// OUTPUT vesrion fichier sur disque dur
$spreadsheet->getProperties()
    ->setTitle('Statistiques hebdomadaires')
    ->setSubject('Statistiques hebdomadaires')
    ->setDescription('Statistiques ventes hebdomadaires')
    ->setCreator('A.R.T.')
    ->setLastModifiedBy('A.R.T.');
$writer = new Xlsx($spreadsheet);
//$writer->save("/var/www/html/dolibarrdelta/documents/ecm/rapports/Stats_inter_". gmdate('D, d M Y H:i:s').".xlsx");
//$writer->save("Stats_inter_". gmdate('D, d M Y H:i:s').".xlsx");
$writer->save(gmdate('Ymd')."_Stats_comm_.xlsx");
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

