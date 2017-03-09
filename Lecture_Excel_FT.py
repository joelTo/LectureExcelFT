import xlrd

wb_reference = xlrd.open_workbook('QueryResult.xlsx')
wb_comparaison = xlrd.open_workbook('FT_MODE_20170307.xlsx')

Feuille_reference = wb_reference.sheet_names()
Feuille_comparaison = wb_comparaison.sheet_names()


sh=wb_reference.sheet_by_name(u'IBM Rational ClearQuest Web')
#recupere la colonne Id du tableau de reference tout doit sortie de ClearQuest
coloneId = sh.col_values(0)

#recupere la colonne libelle du tableau de reference tout doit sortie de ClearQuest
coloneLibelle= sh.col_values(1)

#recupere la colonne State du tableau de reference tout doit sortie de ClearQuest
coloneState=sh.col_values(2)

#recupere la colonne  de la gravte de reference tout doit sortie de ClearQuest
coloneGravite=sh.col_values(3)

print coloneId[2]

#for ligne in coloneId:
#    print "coloneId :",ligne
