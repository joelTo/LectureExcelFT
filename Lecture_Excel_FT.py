import xlrd

wb_reference = xlrd.open_workbook('QueryResult.xlsx')
wb_comparaison = xlrd.open_workbook('FT_MODE_20170307.xlsx')

Feuille_reference = wb_reference.sheet_names()
Feuille_comparaison = wb_comparaison.sheet_names()

#--------------------INPUT Reference -------------------------

sh=wb_reference.sheet_by_name(u'IBM Rational ClearQuest Web')
#recupere la colonne Id du tableau de reference tout doit sortie de ClearQuest
coloneId = sh.col_values(0)

#recupere la colonne libelle du tableau de reference tout doit sortie de ClearQuest
coloneLibelle= sh.col_values(1)

#recupere la colonne State du tableau de reference tout doit sortie de ClearQuest
coloneState=sh.col_values(2)

#recupere la colonne  de la gravte de reference tout doit sortie de ClearQuest
coloneGravite=sh.col_values(3)


inputExcelReference={}
for idx,ligne in enumerate(coloneId):
    inputExcelReference[ligne]={coloneLibelle[idx],coloneState[idx],coloneGravite[idx]}

#----------------------
sh=wb_comparaison.sheet_by_name(u'IBM Rational ClearQuest Web')
#recupere la colonne Id du tableau de reference tout doit sortie de ClearQuest
coloneId_comparaison = sh.col_values(0)

#recupere la colonne libelle du tableau de reference tout doit sortie de ClearQuest
coloneLibelle_comparaison= sh.col_values(1)

#recupere la colonne State du tableau de reference tout doit sortie de ClearQuest
coloneState_comparaison=sh.col_values(2)

#recupere la colonne  de la gravte de reference tout doit sortie de ClearQuest
coloneGravite_comparaison=sh.col_values(3)


inputExcelComparaison={}
for idx,ligne in enumerate(coloneId_comparaison):
    inputExcelComparaison[ligne]={coloneLibelle_comparaison[idx],coloneState_comparaison[idx],coloneGravite_comparaison[idx]}
 #---------------------  

print inputExcelReference['CNAM_00324999']
if inputExcelComparaison['CNAM_00324999']:
    print inputExcelComparaison['CNAM_00324999']

 print 'je suis passe'  