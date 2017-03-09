import xlrd

wb_reference = xlrd.open_workbook('QueryResult.xlsx')
wb_comparaison = xlrd.open_workbook('FT_MODE_20170307.xlsx')

Feuille_reference = wb_reference.sheet_names()
Feuille_comparaison = wb_comparaison.sheet_names()


sh=wb_reference.sheet_by_name(u'IBM Rational ClearQuest Web')
colonne1 = sh.col_values(0)

sh=wb_comparaison.sheet_by_name(u'IBM Rational ClearQuest Web')
colonne2 = sh.col_values(1)


print "Nbr de colonne dans le fichier de reference de ClearQuest est de : "
print len(colonne1)

print "Nbr de colonne dans notre Excel est de : "
print len(colonne2)

r1,r2= set(colonne1),set(colonne2)
print "il manque les FT suivantes dans le "
print r1.difference(r2)

print len(colonne2)
print len(r2)

if len(colonne2)==len(r2):
    print "Il n'y a pas de doublon dans le fichier excel"
else:
    print "Il y a des doublons et ce sont les suivantes :"
    print r2.difference(colonne2)