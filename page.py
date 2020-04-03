# python -m pip install xlrd
from string import Template
import xlrd

f = open("./html/annuaire.html", "w",  encoding='utf8')
annuaire = "<h1>Annuaire</h1>\n" 
f.write(annuaire)
f.close()

book = xlrd.open_workbook('pages.xlsx', encoding_override="utf8")
sheet=book.sheet_by_index(0)

num_rows=sheet.nrows
num_col=sheet.ncols
print(num_col, num_rows)

for i in range (2,num_rows-1) :
    d = {   'nom': sheet.cell(i,0).value,
        'prenom': sheet.cell(i,1).value,
        'trigramme': sheet.cell(i,2).value,
        'email': sheet.cell(i,3).value,
        'photo': sheet.cell(i,4).value,
        'bureau': sheet.cell(i,5).value,
        'fonction': sheet.cell(i,6).value.replace('\n', '</li><li>'),
        'unite': sheet.cell(i,7).value,
        'responsabilite': sheet.cell(i,8).value.replace('\n', '</li><li>'),
        'diplome': sheet.cell(i,9).value.replace('\n', '</li><li>'),        
        'linkedin': sheet.cell(i,10).value
    }

    if ( d['fonction'] ) :
        if (d['linkedin']):
                d['linkedin'] = '<div class="linkedin"><a href="'  + d['linkedin'] + '"><img src="linkedin.png" alt="linkedin" width="30" height="30"></a></div>'
        if (d['responsabilite']):
            d['responsabilite'] = "<ul><li>" + d['responsabilite'] + "</li></ul>"
        if (d['unite']) :
            d['unite'] = '(' + d['unite'] + ')'
        f = open("template.html", mode="r",  encoding='utf8')
        src = Template(f.read())
        f.close()
        result = src.substitute(d)

        f = open("./html/annuaire.html", "a",  encoding='utf8')
        annuaire = "<a href='https://www.ecam.be/annuaire/" + d['trigramme'] + ".html' >" + d['prenom'] + " " + d['nom'] + "</a>\n" 
        f.write(annuaire)
        f.close()

        f = open("./html/" + d['trigramme'] + ".html", "w",  encoding='utf8')
        f.write(result)
        f.close()