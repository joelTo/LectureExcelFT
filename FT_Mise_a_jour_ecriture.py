import xlrd
import xlwt
from xlutils.copy import copy
import os
import time

#initialisaition des variables
fichier_reference='QueryResult.xlsx'
fichier_a_comparer='FT_MODE_20170307.xlsx'
fichier_de_sortie='output.xls'
fichier_compte_rendu="compte_rendu.xls"
date = time.strftime('%d/%m/%y %H:%M',time.localtime()) 

#variables:
num_ligne_compte_rendu=0


wb_reference = xlrd.open_workbook(fichier_reference)
wb_comparaison = xlrd.open_workbook(fichier_a_comparer)

if os.path.isfile(fichier_compte_rendu): 
    wb_compte_rendu=xlrd.open_workbook(fichier_compte_rendu)
    feuille_compte_rendu=wb_compte_rendu.sheet_by_name(u'compteRenduFT')
    num_ligne_compte_rendu=len(feuille_compte_rendu.col_values(0))  
    wb_compte_rendu = copy(wb_compte_rendu)
    excel_compte_rendu=wb_compte_rendu.get_sheet(0)
    print num_ligne_compte_rendu;

else : 
    wb_compte_rendu=xlwt.Workbook()
    excel_compte_rendu = wb_compte_rendu.add_sheet('compteRenduFT')
    excel_compte_rendu.write(0,0,"Mise a jour le :")
    excel_compte_rendu.write(0,1,date)
    #Type : ajout - modif 
    excel_compte_rendu.write(1,0,'Date')
    excel_compte_rendu.write(1,1,'numero_FT')
    excel_compte_rendu.write(1,2,'type_de_modification')
    excel_compte_rendu.write(1,3,'etat_precedent')
    excel_compte_rendu.write(1,4,'etat_actuellement')
    excel_compte_rendu.write(1,5,'calcul_de_velocite')
    num_ligne_compte_rendu=2
    
    

Feuille_reference = wb_reference.sheet_names()
Feuille_comparaison = wb_comparaison.sheet_names()

#couleur des lignes modifiees
style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on')

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


test = xlrd.open_workbook(fichier_a_comparer)
test.sheet_by_index(0).cell(0,0).value
wb = copy(test)



#Valeur de debut nouvelle ecriture
Index_dernierElement=len(r2)

#creation 
for idx,item in enumerate(ListeFtManquante):
    print item,inputExcelReferenceParCategorie[item]
    print idx+Index_dernierElement
    #ecriture du numero FT
    wb.get_sheet(0).write(idx+Index_dernierElement-1,1,item,style0)
    #ecriture du libelle
    wb.get_sheet(0).write(idx+Index_dernierElement-1,2,inputExcelReferenceParCategorie[item]['libelle'],style0)
    #ecriture du state
    wb.get_sheet(0).write(idx+Index_dernierElement-1,3,inputExcelReferenceParCategorie[item]['state'],style0)
    #ecriture de la gravite
    wb.get_sheet(0).write(idx+Index_dernierElement-1,4,inputExcelReferenceParCategorie[item]['gravite'],style0)
    #ecriture de la dateCreation
    wb.get_sheet(0).write(idx+Index_dernierElement-1,5,inputExcelReferenceParCategorie[item]['dateCreation'],style0)
    #ecriture de la dateModif
    wb.get_sheet(0).write(idx+Index_dernierElement-1,6,inputExcelReferenceParCategorie[item]['dateModif'],style0)
    #ecriture de la acp
    wb.get_sheet(0).write(idx+Index_dernierElement-1,7,inputExcelReferenceParCategorie[item]['acp'],style0)

    #ajout compte rendu
    print type(excel_compte_rendu) 
    excel_compte_rendu.write(num_ligne_compte_rendu,0,date)
    excel_compte_rendu.write(num_ligne_compte_rendu,1,item)
    excel_compte_rendu.write(num_ligne_compte_rendu,2,'ajout')
    num_ligne_compte_rendu=num_ligne_compte_rendu+1


  
    

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
                print colonneId_comparaison.index(FT)+1 
                print (inputExcelReferenceParCategorie[FT])['gravite']
                wb.get_sheet(0).write(colonneId_comparaison.index(FT),4,(inputExcelReferenceParCategorie[FT])['gravite'],style0)
                nbFT=True


         if ((inputExcelReferenceParCategorie[FT])['state'] != (inputExcelComparaisonParCategorie[FT]['state'])):
            if (nbFT==False):
                print "---------------------------------------------------------------"
                print "-----------------------------",nbrModif,") ",FT,"----------------------------------" 
            print 'reference' ,(inputExcelReferenceParCategorie[FT])['state'] 
            print 'comparaison', (inputExcelComparaisonParCategorie[FT])['state'] 
            print colonneId_comparaison.index(FT)+1
            wb.get_sheet(0).write(colonneId_comparaison.index(FT),3,(inputExcelReferenceParCategorie[FT])['state'],style0)
            print (inputExcelReferenceParCategorie[FT])['state']           
            nbFT=True

            #Ajout modification
            excel_compte_rendu.write(num_ligne_compte_rendu,0,date)
            excel_compte_rendu.write(num_ligne_compte_rendu,1,FT)
            excel_compte_rendu.write(num_ligne_compte_rendu,2,'modif')
            excel_compte_rendu.write(num_ligne_compte_rendu,3,(inputExcelComparaisonParCategorie[FT])['state'])
            excel_compte_rendu.write(num_ligne_compte_rendu,4,(inputExcelReferenceParCategorie[FT])['state'])
            num_ligne_compte_rendu=num_ligne_compte_rendu+1
      
wb.save(fichier_de_sortie)
wb_compte_rendu.save('compte_rendu.xls')                   



    
