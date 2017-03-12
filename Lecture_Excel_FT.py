import xlrd

wb_reference = xlrd.open_workbook('QueryResult.xlsx')
wb_comparaison = xlrd.open_workbook('FT_MODE_20170307.xlsx')

Feuille_reference = wb_reference.sheet_names()
Feuille_comparaison = wb_comparaison.sheet_names()

#--------------------INPUT Reference -------------------------

sh=wb_reference.sheet_by_name(u'IBM Rational ClearQuest Web')
#recupere la colonne Id du tableau de reference tout doit sortie de ClearQuest
coloneId_reference = sh.col_values(0)

#recupere la colonne libelle du tableau de reference tout doit sortie de ClearQuest
coloneLibelle_reference= sh.col_values(1)

#recupere la colonne State du tableau de reference tout doit sortie de ClearQuest
coloneState_reference=sh.col_values(2)

#recupere la colonne  de la gravte de reference tout doit sortie de ClearQuest
coloneGravite_reference=sh.col_values(3)


inputExcelReference={}
inputExcelReferenceParCategorie={}
for idx,ligne in enumerate(coloneId_reference):
    inputExcelReference[ligne]={coloneLibelle_reference[idx],coloneState_reference[idx],coloneGravite_reference[idx]}
    inputExcelReferenceParCategorie[ligne]={'libelle':coloneLibelle_reference[idx],'state':coloneState_reference[idx],'gravite':coloneGravite_reference[idx]}

#----------------------
sh=wb_comparaison.sheet_by_name(u'IBM Rational ClearQuest Web')
#recupere la colonne Id du tableau de reference tout doit sortie de ClearQuest
coloneId_comparaison = sh.col_values(1)

#recupere la colonne libelle du tableau de reference tout doit sortie de ClearQuest
coloneLibelle_comparaison= sh.col_values(2)

#recupere la colonne State du tableau de reference tout doit sortie de ClearQuest
coloneState_comparaison=sh.col_values(3)

#recupere la colonne  de la gravte de reference tout doit sortie de ClearQuest
coloneGravite_comparaison=sh.col_values(4)


inputExcelComparaison={}
inputExcelComparaisonParCategorie={}
for idx,ligne in enumerate(coloneId_comparaison):
    inputExcelComparaison[ligne]={coloneLibelle_comparaison[idx],coloneState_comparaison[idx],coloneGravite_comparaison[idx]}
    inputExcelComparaisonParCategorie[ligne]={'libelle':coloneLibelle_comparaison[idx],'state':coloneState_comparaison[idx],'gravite':coloneGravite_comparaison[idx]}
    
 #---------------------  

#Affiche les FT manquantes
r1,r2= set(coloneId_reference),set(coloneId_comparaison)
print "---------------------------------------------------------------"
print "---------------------------------------------------------------"
print "-   ","Nbr de FT manquante dans le document est de : ", len(r1.difference(r2))
print "-   ",r1.difference(r2)
print "---------------------------------------------------------------"
print "---------------------------------------------------------------"


listeFtReference= inputExcelReference.keys()
nbrModif=0;
for FT in listeFtReference:
      # Si cette cle est trouve alors rentrer et faire la comparaison entre la FT des documents. 
      if inputExcelComparaison.has_key(FT):
         if(inputExcelReference[FT].difference(inputExcelComparaison[FT])):
             nbrModif= nbrModif + 1;
             nbFT = False
             if ((inputExcelReferenceParCategorie[FT])['gravite'] != (inputExcelComparaisonParCategorie[FT]['gravite'])):
                 print "---------------------------------------------------------------"
                 print "-----------------------------",nbrModif,") ",FT,"----------------------------------" 
                 print 'reference' ,(inputExcelReferenceParCategorie[FT])['gravite'] 
                 print 'comparaison', (inputExcelComparaisonParCategorie[FT])['gravite']
                 nbFT=True

             if ((inputExcelReferenceParCategorie[FT])['state'] != (inputExcelComparaisonParCategorie[FT]['state'])):
                if (nbFT==True):
                    print "---------------------------------------------------------------"
                    print "-----------------------------",nbrModif,") ",FT,"----------------------------------" 
                print 'reference' ,(inputExcelReferenceParCategorie[FT])['state'] 
                print 'comparaison', (inputExcelComparaisonParCategorie[FT])['state']
                nbFT=True
             
             if ((inputExcelReferenceParCategorie[FT])['state'] != (inputExcelComparaisonParCategorie[FT]['state'])):
                if (nbFT==True):
                    print "---------------------------------------------------------------"
                    print "-----------------------------",nbrModif,") ",FT,"----------------------------------" 
                print 'reference' ,(inputExcelReferenceParCategorie[FT])['state'] 
                print 'comparaison', (inputExcelComparaisonParCategorie[FT])['state']
                nbFT=True

                 



    