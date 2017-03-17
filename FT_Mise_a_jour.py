import xlrd
import xlwt

fichier_reference='QueryResult.xlsx'
fichier_a_comparer='FT_MODE_20170307.xlsx'

wb_reference = xlrd.open_workbook(fichier_reference)
wb_comparaison = xlrd.open_workbook(fichier_a_comparer)

Feuille_reference = wb_reference.sheet_names()
Feuille_comparaison = wb_comparaison.sheet_names()

#--------------------INPUT Reference -------------------------

sh=wb_reference.sheet_by_name(u'IBM Rational ClearQuest Web')
#recupere la colonne Id du tableau de reference tout doit sortie de ClearQuest
colonneId_reference = sh.col_values(0)

#recupere la colonne libelle du tableau de reference tout doit sortie de ClearQuest
colonneLibelle_reference= sh.col_values(1)

#recupere la colonne State du tableau de reference tout doit sortie de ClearQuest
colonneState_reference=sh.col_values(2)

#recupere la colonne  de la gravte de reference tout doit sortie de ClearQuest
colonneGravite_reference=sh.col_values(3)

#recupere la date de creation du document de reference 
colonneDateCreation_reference=sh.col_values(4)

#recupere la DateDerniereModif du document de Reference
colonneDateModif_reference=sh.col_values(5)

#recupere le code de la FT du document de Reference
colonneACP_reference=sh.col_values(6)



inputExcelReference={}
inputExcelReferenceParCategorie={}
for idx,ligne in enumerate(colonneId_reference):
    inputExcelReference[ligne]={colonneLibelle_reference[idx],colonneState_reference[idx],colonneGravite_reference[idx],colonneDateCreation_reference[idx],colonneDateModif_reference[idx],colonneACP_reference[idx]}
    inputExcelReferenceParCategorie[ligne]={'libelle':colonneLibelle_reference[idx],'state':colonneState_reference[idx],'gravite':colonneGravite_reference[idx],'dateCreation':colonneDateCreation_reference[idx],'dateModif':colonneDateModif_reference[idx],'acp':colonneACP_reference[idx]}

#----------------------
sh=wb_comparaison.sheet_by_name(u'IBM Rational ClearQuest Web')
#recupere la colonne Id du tableau de reference tout doit sortie de ClearQuest
colonneId_comparaison = sh.col_values(1)

#recupere la colonne libelle du tableau de reference tout doit sortie de ClearQuest
colonneLibelle_comparaison= sh.col_values(2)

#recupere la colonne State du tableau de reference tout doit sortie de ClearQuest
colonneState_comparaison=sh.col_values(3)

#recupere la colonne  de la gravte de reference tout doit sortie de ClearQuest
colonneGravite_comparaison=sh.col_values(4)




inputExcelComparaison={}
inputExcelComparaisonParCategorie={}
for idx,ligne in enumerate(colonneId_comparaison):
    inputExcelComparaison[ligne]={colonneLibelle_comparaison[idx],colonneState_comparaison[idx],colonneGravite_comparaison[idx]}
    inputExcelComparaisonParCategorie[ligne]={'libelle':colonneLibelle_comparaison[idx],'state':colonneState_comparaison[idx],'gravite':colonneGravite_comparaison[idx]}
   
 #---------------------  

#Affiche les FT manquantes
r1,r2= set(colonneId_reference),set(colonneId_comparaison)
print "---------------------------------------------------------------"
print "---------------------------------------------------------------"
print "-   ","Nbr de FT manquante dans le document est de : ", len(r1.difference(r2))
ListeFtManquante=r1.difference(r2)
print "-   ",ListeFtManquante
print "---------------------------------------------------------------"
print "---------------------------------------------------------------"

#Affichage + rajout dans le futur fichier l'ecrite à la dernière place de cette FT
for item in ListeFtManquante:
    print item,inputExcelReference[item]
    

listeFtReference= inputExcelReference.keys()
nbrModif=0
for FT in listeFtReference:
      # Si cette cle est trouve alors rentrer et faire la comparaison entre la FT des documents. 
      if inputExcelComparaison.has_key(FT):
         nbFT=False
         nbrModif=nbrModif+1
         if ((inputExcelReferenceParCategorie[FT])['gravite'] != (inputExcelComparaisonParCategorie[FT]['gravite'])):
                print "---------------------------------------------------------------"
                print "-----------------------------",nbrModif,") ",FT,"----------------------------------" 
                print 'reference' ,(inputExcelReferenceParCategorie[FT])['gravite'] 
                print 'comparaison', (inputExcelComparaisonParCategorie[FT])['gravite']
                nbFT=True

         if ((inputExcelReferenceParCategorie[FT])['state'] != (inputExcelComparaisonParCategorie[FT]['state'])):
            if (nbFT==False):
                print "---------------------------------------------------------------"
                print "-----------------------------",nbrModif,") ",FT,"----------------------------------" 
            print 'reference' ,(inputExcelReferenceParCategorie[FT])['state'] 
            print 'comparaison', (inputExcelComparaisonParCategorie[FT])['state']
            nbFT=True
        

                 



    