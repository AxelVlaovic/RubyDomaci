require './biblioteka.rb'

t1 = Table.new './Table1.xlsx'
t2 = Table.new './Table2.xls'
t3 = Table.new './Table3.xlsx'

puts "--------------------------"
#Parsiranje xlsx i xls tabela , mergovana polja , ignorisanje praznih redova , ignorisanje redova sa: total ili subtotal

p t1.table
p t2.table
p t3.table 

puts "--------------------------"
 #Pristup tabeli kao dvodimenzionalnom nizu

p t1.table[0][2]
p t1.table[3][0]
p t2.table[1][0]
p t3.table[4][1]

puts "--------------------------"
 #Pristup redovima preko t.row,kao i pojedinacnim elementima u redu

p t1.row(2)
p t1.row(3)[0]
p t3.row(4)
p t2.row(0)[0]

puts "--------------------------"
 #Iteriranje kroz sve ćelije tabele pomoću implementacije Enumerable

t1.each do |cell|
    puts cell
end

puts "|------|"

t2.each do |cell|
    puts cell
end

puts "--------------------------"
 #[] sintaksa,dobijanje cijele kolone na upit t["headerExample"],pristup pojedinacnim elementima te kolone t["he"][0]

p t1["prvaKolona"]
p t1["trecaKolona"][0]
p t3["drugaKolona"][1]
p t2["k1"]

puts "--------------------------"
 #Direktan pristup kolonama preko istoimenih metoda
 p t1.prvaKolona
 p t1.prvaKolona[0]
 p t2.k2
 p t3.trecaKolona

 puts "--------------------------"
 #Računanje sume neke kolone preko istoimenih metoda

p t1.drugaKolona.sum
p t1.trecaKolona.sum
p t2.k2.sum
p t3.trecaKolona.sum

puts "--------------------------"
 #NIJE IMPLEMENTIRANO|| Imena metoda su generisana ali one ne vraćaju odgovarajući red 
p t1.prvaKolona.lubenica
p t3.prvaKolona.kivi

puts "--------------------------"
 #Sabiranje dvije tabele

p t1+t2
new_table = t1+t3

t4 = Table.new new_table
p t4.table
p t4.row(5)
p t4.prvaKolona

puts "--------------------------"
 #Oduzimanje dvije tabele

modified_table = t4-t3

t5 = Table.new modified_table

p t5.table
p t5.row(2)
p t5.drugaKolona[0]







