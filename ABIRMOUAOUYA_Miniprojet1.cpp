/****----------------------    MINI PROJET : MANIPULATION D'UN FICHIER EXCEL PAR UN PROGRAMME C++   --------------------------****/


/*----------------------------------------------------------------------------------------------------------------------------------
####################################################################################################################################

Organisation    : SICoM - FST 
Mini projet  	: MANIPULATION D'UN FICHIER EXCEL PAR UN PROGRAMME C++
Nom du projet 	: AbirMouaouya_miniprojet1.dev
Nom de fichier 	: AbirMouaouya_miniprojet1.cpp
Description 	: ce programme permet de manipuler un fichier EXCEL , qui stocke les informations des étudiants(CNE,NOM,PRENOM,LES NOTES)   

fonctions utilisées :   
       	* AjouterEtudiant() , consiste à ajouter un étudiant au fichier excel .
		* SupprimerEtudiant,consiste à supprimer un étudiant du fichier excel .
		* ModifierNote (),modifier la note d'un utilsateur donné.
		* AfficherNote(),afficher les resultat d'un étudiant donné  .
		* AfficherListe(),apres l'ajout des etudiants on affiche toute la liste , sinon l'utilisateur peut afficher toute la iste avant de modifier les notes d'un etudiant 
		* EtudiantExiste(), verifie si le cne saisie existe dans la liste des étudiants 
		* Fonctions predefinies sur les directives iostream,fstream,string ,string.h 
		
les extensions du fichier Excel valident sur ce programme ,sont : xls , xlss   ( depend de séparateur \t ) 
Historique 		: 10/03/2021 
COPYRIGHT 		: FST-USMBA 
Crée par 		: ABIR MOUAOUYA 
Encadré par 	: HICHAM GHENNIOUI 

#####################################################################################################################################
-----------------------------------------------------------------------------------------------------------------------------------*/


	   

 /*____________________________ les directives de  preprocesseur ________________________________________________*/
/*--------------------------------------------------------------------------------------------------------------*/


#include<locale>            // directive de  preprocesseur qui gére les extensions d'affichage accenté  
#include<iostream>         // directive de  preprocesseur qui gére les extensions des entrées/sorties 
#include<fstream>	   	  // directive de preprocesseur chargé des extensions de manipulation des fichiers 
#include<string.h>       // directive de  preprocesseur qui gére les extensions des fonctions de maniplulation des chaines de caractéres 
#include<string>        // directive de  preprocesseur qui gére les extensions les chaines de caractéres 
using namespace std;   // espace de définition des commandes pour éviter la répitition de std 



/****_____________________________   Declaration de la Structure   _____________________________****/
/***-----------------------------------------------------------------------------------------***/


/*  LA structure etudiant : a comme argument  les informations d'un etudiant,chaque l'etudiant à un CNE qui sera l'id .
dont  on va charger les informations saisies au clavier afin de les insérer dans notre base de donnée à base d'un fichier EXCEL */
struct etudiant{
	  char nom[40];
	  char prenom[40];
	  char CNE[10];          // ce champs sera l'id de chaque eleve pour faciliter le faite d'exrtaire les informations
	  double  noteCC;       //note de controle continue 
	  double  noteTP;	   //note de travaux pratique
	  double  noteCU;     //note de controle unifié
	  string resultat;   // rattrapage, validé ou non validé .	  
};



 /****________________________   Prototype des fonctions  _____________________________________****/
/***-------------------------------------------------------------------------------------------***/
etudiant SaisirEtudiant();                             // return un etudiant saisie au clavier                                
void AjouterEtudiant(char * fileName);                // consiste à ajouter un étudiant au fichier excel .
void SupprimerEtudiant(char * fileName);    	     // consiste à supprimer un étudiant du fichier excel .
void ModifierNote (char * fileName);       		    // modifier la note d'un étudiant  donné par son CNE .
void AfficherEtudiant(char* cne,char * fileName);  // afficher  d'un étudiant donné par son CNE  .
void AfficherNote(char * fileName);               // afficher les notes  d'un étudiant donné par son CNE  .
void filecopy(char* fileName) ;                  // copier un fichier .
void AfficherListe(char * fileName);      	    // afficher toute la liste des etudiants .
bool EtudiantExiste(char* cne,char * fileName);// verifie si le cne saisie est dans la liste.



 /***_____________________________________   Fonction Principale   __________________________________________***/
/*------------------------------------------------------------------------------------------------------------*/
int main(void) {  
    // debut de programme 
    setlocale(LC_CTYPE, "fra");  // corrigé l'affichage des accents 
    
    int choix; 		      // la fonction choisée par l'utilisateur
    int continu=0;       // la variable de décision pour que l'utilisateur choisi s'il veut continuer ou arreter
	int  n;        	    // taille du nom de fichier que l'utilisateur va saisir  
    char  fileName[30];// le nom de fichier 
    string test;      // pour tester les 4 caracteres s'il sont == ".xls" 
    
    
    
   cout << "********************************** Bienvenue à notre programme de manipulation des notes des étudiants **************************************\n\n"<<endl; 
   cout <<endl; 

	 
	 
//-----------------------------***  demander à l'utilisateur le nom du  fichier qu'on va manipuler ***-----------------------------
   
    cout <<" saisir le nom de vorte fichier sous format nom_fichier.xls" <<endl;
    cin  >> fileName;
    cout <<endl<<endl<<endl;


  do{     
//-------------------------------------------**** le menu ****----------------------------------------------------------------------
    cout << " ||  menu principal  ________________________________________ "<<endl;
    cout << " ||                                                          |"<<endl;
    cout << " || 1) Ajouter un Etudiant .                                 |"<<endl;
    cout << " || 2) Supprimer un Etudiant .                               |"<<endl;
    cout << " || 3) Modifier les données d'un Etudiant .                  |"<<endl;
    cout << " || 4) Afficher les notes d'un Etudiant .                    |"<<endl;
    cout << " || 5) Afficher la liste des  Etudiants .                    |"<<endl;
    cout << " || 6) Quitter .                                             |"<<endl;
    cout << " ||__________________________________________________________|"<<endl;
    cin  >> choix;      
    
   // system("cls");   // supprimer tout de ce qui est écrit à l'écran pour que le reste sera clair 
   
    
/**---------------------------- manipulation des choix qui vont de 1 à 6 ------------------------------**/
    switch(choix)
    { 
     case 1: //Ajouter Etudiant
             AjouterEtudiant(fileName);
             break;
     case 2: //Supprimer un Etudiant . 
             SupprimerEtudiant(fileName);
     	    break;
     case 3://Modifier les données d'un Etudiant .
             ModifierNote(fileName);
     	    break;
     case 4://Afficher les notes d'un Etudiant .
	         AfficherNote(fileName) ;
     	    break;
     case 5://Afficher la liste des  Etudiants . 
             AfficherListe(fileName);
     	    break;
     case 6 : return 0;
              break;
     default : cout << "Merci de saisir un des nombres indiqués dans le menu \n "<<endl;               
	}	
	
	cout <<endl<<endl<<endl;
	
	//	demander à l'utilisateur s'il faut continuer  .
	cout <<" voulez vous : \n 0* quitter \n 1* retournez au menu principale "<<endl;
    cin >> continu;
   
}while(continu!=0);

/*----------------------------------------     fin  de fonction principale      --------------------------------------------*/ 
}




/****________________________________________________ La declaration du corps des fonctions  _____________________________________________________****/
/*   ---------------------------------------------------------------------------------------------------------------------------------------------  */
 
 //fonction saisir etudiant permet de remplir les champs d'une varible structuré de type etudiant par des variables saisies au clavier afin de la  retourner
// cette variable est utilisée en insertion dans  le fichier .
 etudiant SaisirEtudiant() {
        etudiant E;        
        cout <<" ***	saisir les informations de l'etudiant "<< endl;
        
        do{
		cout <<"   *CNE       :"<<endl;
         cin >> E.CNE;
		}while(strlen(E.CNE)!=10); //test de saisie sur le CNE .
        cout <<"   *Nom        :"<<endl;
         cin >> E.nom;
        cout <<"   *PreNom     :"<<endl;
         cin >> E.prenom;
         do{
        cout <<"   *note de CC :"<<endl;
         cin >> E.noteCC;
		 }while(E.noteCC<0 || E.noteCC>20);
		 do{
        cout <<"   *note de TP :"<<endl;
         cin >> E.noteTP;
          }while(E.noteTP<0 || E.noteTP>20);
         do{
        cout <<"   *note de CU :"<<endl;
         cin >> E.noteCU; 
		  }while(E.noteCU<0 || E.noteCU>20); 
         
         //calcul du resultat final 
		double note;
        note=(E.noteTP*0.25)+(E.noteCC*0.25)+(E.noteCU*0.5);
        if(note>=12)
              E.resultat="v";           //validé 
        else // note <12 
             if(note>=5)
                E.resultat="R";         //RATT
             else 
                E.resultat="NV"  ;    // non validé  
                
	   return E;
	   }   
	   
	   
   /*fonction d'ajout des etudiants prend comme argument le fichier à manipuler ,elle demande le nombre des etudiants à saisir
on ouvre le fichier de façon que son contenu ne sera pas ecrasé pour ne perder pas data des etudiants,
avec le mode ios::app , et on ajoute les etudiants par la fonction saisir etudiant , 
en fin  on affiche  un  message de réussite d'ajout . */
void  AjouterEtudiant(char* fileName)
     {   
      int nbrEtudiant,i=0;
      etudiant etd;
	   
	  do{ //on demande le nombre des etudiants à saisie 
	      cout << "	quel est  le nombre des etudiants que voulez vouz ajouter ?" <<endl;
          cin  >> nbrEtudiant;  
	   }while(nbrEtudiant<=0); //controle de saisie 
	     
	  cout <<endl<<endl<<endl;
	  
      ofstream file (fileName,ios::app); //ios::app pour n'ecraser pas le fichier s'il exsite 
      
      if(!file.is_open())    //tester  l'ouverture
        	cout << "Echec d'ouverture de fichier"<<endl; //message d'erreur si jamais on a diff d'ouvrire le fichier
	  else{ 
	       // boucle pour saisir les informations de chaque etudiant 
	       do 
		   {
             	etd=SaisirEtudiant();
		    	//apres avoir les informations de chaque etudiant on va les stockées dans notre fichier
		    	if(EtudiantExiste(etd.CNE,fileName))
		    	   cout<<endl<<"Etudiant existe  deja ! "<<endl;
		    	else{
                file <<etd.CNE<<"	"<<etd.nom<<"	"<<etd.prenom<<"	"<<etd.noteCC<<"	"<<etd.noteTP<<"	"<<etd.noteCU<<"	"<<etd.resultat <<endl;	
                i++;
				}
		    // le séparateur pour l'extension .xls est  tabulation 
		   }while(i<nbrEtudiant);
		   cout <<endl<<endl<<endl;	
		   	   
		   cout<<"la saisie est faite avec succée"<<endl; // message de confirmation d'ajout
		 }
	file.close();   // fermer le fichier 
 } 
     	
 
 
/* on se basons sur le CNE comme identifiant de l'etudiant , on teste  les 10 premiers caracteres de chaque ligne , on cherche si ce cne existe ou non 
 c'est une fonction de type boolen retourne vrai si le cne existe et faux sinon .*/
bool EtudiantExiste(char* cne,char * fileName){
	
	 string ligne;
	 
	 ifstream file(fileName); // ouvrir le fichier on mode in
	  
	 if(file.is_open())      //tester si l'ouverture est bien faite 
	    while(getline(file,ligne)) // charger la ligne de fichier dans la variable ligne 
	      if(cne==ligne.substr(0,10)) //substr(0,10) permet de recupere de 0 au 10 éme caractere de la ligne,afin de le comparer avec le cne donné en arguments 
	         return true ;  
	 return false;
}



/* fonction supprime un etudiant donné par son CNE  : 	  
cette fonction à comme principe : demander le cne de l'etudiant à supprimer , apres on ouvre un autre fichier intermidiare ,dont 
on va charger ligne par ligne jusqu'on arrive à la ligne qu'on veut supprimer , on va la depasser sans la copie 
apres on va recopier les lignes du fichier intermidiare dans le fichier principale qu'on ouvre en mode out sans app pour ecraser
les lignes anciennes apres on supprime le fichier intermidiare , 
cette fonction fait appel à la fonction filecopy  qui est chargée de copie un fichier dans un autre et supprimer le fichier intermidiare*/
void SupprimerEtudiant(char * fileName)
     { char cne[10]; 
	   string ligne;               // dans on va recuperer les ligne de fichier 
	   int confirmation=0;
	   
       ifstream file(fileName);   //ouvrir le fichier principale en mode in pour extraire les information   
       ofstream help("help.xls"); //cree un autre fichier, dont on va stocker  les lignes qu'on va garder, a fin de le recopie dans le fichieer principal
       cout << "saisir le cne de l'etudiant à supprimer"<<endl;
       cin  >> cne;
       
       if(file.is_open() && help.is_open() ){
         if(EtudiantExiste(cne,fileName)){ // verifie l'existance de  l'etudiant avant de supprimer
      	     
	      	 //afficher  l'etudiant à supprimer , pour eviter les erreurs si jamais l'utilisateur à saisie un cne d'un autre etudiant  
	      	 cout <<endl; 
			 AfficherEtudiant(cne,fileName);     	
	      	 // demander la confirmation de supprission 
	      	cout<<endl;
	      	cout<<"	voulez-vous vraiment suuprimer cet etudiant  "<<endl<<"	1* oui 		0*non "<<endl;
	      	cin >> confirmation;
      	
	      	if( confirmation==1){
	           while( getline(file,ligne)){  // charger les lignes du fichier 
		          if(cne==ligne.substr(0,10))//chercher la ligne qui contien le cne saisie 
		              continue;              //lorsque on arrive au etudiant mentionné on saute la ligne sans la copier
		          else                    
		            help << ligne<<endl;  // copie les lignes qu'on ne contient par le cne donné  dans le fichier help 
					  }        
		        cout <<endl<<endl<<endl;
				cout<<"suppresion réussie"<<endl;	// message de confirmation de suppression 
			   
		     } 
			else  // si comfirmation =0 on va supprimer le fichier help 
			   remove("help.xls");    
			  } 
		  else 
		    cout <<"cet étudiant n'existe pas "<<endl;}
		    
	 // fermer les deux fichiers   
	file.close();
	help.close();
	if(confirmation==1)// si on a effectué une suppression 
		filecopy(fileName); // copie le fichier help dans le fichier principal	 
}

	  
/* cette fonction sert à recree le fichier principale ,en ecrasant son contenu , afin de l'initialiser par les valeurs
 sollicitée  apres on supprime le fichier help*/
void filecopy(char* fileName){ 
string ligne;
 // reouvrir les deux fichiers 
 
       ofstream file(fileName);   // mode out sans append 
	   ifstream help("help.xls"); // mode in 
	   
	   while( getline(help ,ligne)) // charger les lignes he help 
            file << ligne<<endl; // inserer ses lignes dans notre fichier 
            
	   help.close();	//fermer le fichier intermidiare 
       remove("help.xls");// suoprimer le fichier intermidiare 
       file.close();// fermer le fichier principal 
}	  



//fonction modifie la note d'un etudiant donné par CNE : 
  
/*	apres qu'on  demande le CNE , on le cherche dans notre base de donnee s'il existe on commence notre traitemnt qui est basé
sur le meme principe de  la suppression , en utilisant  un fichier intrmidiaire dont on va copie ligne par ligne jusqu'on arrive à l'etudiant
avec le cne donné , on demande à l'utilisateur de resaisir touts les informations de cet étudiant y compris le cne ,le nom,
et le prenom . mais avant de demander les nouvelles informations on affiche la ligne de l'etudiant à modifier pour faciliter la resaisi  " */ 

void ModifierNote (char* fileName)
     { 
       string ligne;
       char cne[10];
       etudiant etd;  // dans on va charger les nouvelles infos de l'etudiant 
       int confirmation=0;
       
       ifstream file ;		    
	   file.open(fileName); 		// ouvrir le fichier en un mode in 
       ofstream help("help.xls");   // nouveau fichier dont on va garder nos données modifiées 
       
       cout <<endl<<endl<<endl;
       
	   cout << "  saisir le CNE de l'etudiant "<<endl; // demander le cne 
	   cin  >> cne; 
	   
		 // puisque on a pas une methode efficace pour modifer les champs d'une ligne on fait modifier toute la ligne 
	   if(file.is_open() && help.is_open() ){ // verifie si les deux fichiers sont ouverts
         
		     if(!EtudiantExiste(cne,fileName))     // chercher le  CNE dans le fichier principal  
               cout << " etudiant introuvable"<<endl;
            
             else {  // si le cne est trouvé 
                //apres avoir verifier l'existance d'unr  ligne qui contient le cne donné , on l'affiche pour que l'utilisateur verifie les données de l'etudiant 
                AfficherEtudiant(cne,fileName); 	
            
		    	// apres on le demande s'il veut encore le modifier ou non 
				cout <<"	voulez vous modifier cet etudiant "<<endl<<"1* oui 	0*non "<<endl;
				cin  >> confirmation;
	
			    if(confirmation==1) { 
				    // si la reponse est oui , on le demande de saisir les nouvelles informations 
	      			cout <<"  resaisir à nouveau les informations de cet etudiant"<<endl;
	      			etd=SaisirEtudiant(); // on les  charge dans la variable etd de type etudiant
	      
		   			while(getline(file,ligne) )  { // on cherche la ligne  		        
	          			 if(cne==ligne.substr(0,10)) 	// apres la saisie on insère les nouvelles données dans notre fichier  
	           				 help <<etd.CNE<<"	"<<etd.nom<<"	"<<etd.prenom<<"	"<<etd.noteCC<<"	"<<etd.noteTP<<"	"<<etd.noteCU<<"	"<<etd.resultat<<endl;
					  
		                 else  // si le cne ne correspond pas au CNE  donnéé 
		                     help <<ligne<<endl; /* on charge tt simplement  toute la ligne  dans le fichier help*/ }
		            cout<<"modification est bien faite "<<endl; }
		        else 
		           remove("help.xls");
		        } 
		}
		     
	   help.close();
       file.close();
       if(confirmation == 1)
       	 filecopy(fileName); // recopier les infos modifier dans le fichier principal 
	   }
	
	
//afficher un etudiant donnée par son CNE ,et qu'on a deja verifier son existance !																
void AfficherEtudiant( char* cne,char *fileName)
 {
  string ligne;                                 //la variable  dans on va apporter chaque ligne de fichier afin de l'analyser

     ifstream file(fileName);
	 while(getline(file,ligne))
	    if(cne==ligne.substr(0,10))     // on cherche l'apparition du CNE
	       cout <<ligne<<endl;              // ON AFiche la ligne qui le contient 
   	 }
			
			
			
//afficher les notes d'un etudiant donnée par son CNE ,s'il existe !																
void AfficherNote(char *fileName)
 {
  char cne[10];
  string ligne;     //la variable  dans on va apporter chaque ligne de fichier afin de l'analyser
  
  cout << "  saisir le cne de l'etudiant"<<endl;
  cin >>cne;              // on fait nos testes sur le cne puisque il est unique 
  
  if(!EtudiantExiste(cne,fileName)) 
      cout<<"  cne introuvable "<<endl;
  else 
      AfficherEtudiant(cne,fileName);
	  }              // ON AFiche la ligne qui le contient 
   	 

 //afficher tout les etudiants 	
 void AfficherListe(char* fileName)
{
	char mot[40];         // dans on va apporter mot par mot  
	int i=0;             // compteur pour se positionner entre les lignes 
	
    ifstream file (fileName);    // ouvrir le fichier en mode lecture 
    
    if (file.is_open())            // tester si l'ouverture d fichier est bien faite 
    {  // tracer l'en-tete du tableau 
	    cout<<"|   cne          ||   nom       || prenom    || note cc|| note TP||note cu||resultat | "<<endl;
        cout<<"_____________________________________________________________________________________  "<<endl;
        cout<<"-------------------------------------------------------------------------------------  "<<endl;
    
	    while(file >> mot){ //tant que on a pas encore arriver à la fin du fichier on charge le mot actuel dans la variable "mot"
          cout <<"|   "<<mot<<"   |"; // afficher chaque mot entre | | pour bien presenter le tableau  
          i++;     // incrementer le compteur ;
          if(i==7)    // on a 7 informations  pour chaque étudiants ,  
             {i=0;     // apres qu'on affiche toutes les infos d'un etudiant on rend le compteur à 0 
		      cout<<endl;  // un retour à la ligne pour se preparer à afficher les infos de l'etudiant suivant.
 		      cout<<"______________________________________________________________________________________"<<endl; }
	         }
	file.close();}           //fermer le fichier 
	
  else cout << "Impossible d'ouvrir le fichier"<<endl;
 	}


/******************************************************   END ************************************************************************/
/*************************************** copyright : FST FES *************************************************************************/
/******************************** DEVLOPPE PAR : ABIR MOUAOUYA **********************************************************************/ 		
