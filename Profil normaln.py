import numpy as np
import pandas as pd
from pyautocad import Autocad, APoint ,distance 
import math as m 
import random as rm
# Inicjalizacja interfejsu AutoCAD
acad = Autocad()
   
# Odczytanie danych z arkusza Excela 
excel = pd.read_excel('C:/Users/damia/Desktop/Code/Protokół Python/Pyizl/izl.xlsx', sheet_name='Przekroj normalny')

#######
# dane wyjściowe

Rodzaj_drogi = excel['Dane'][1] 
Nawierzchnia = excel['Dane'][3]
Grunt = excel['Dane'][5]
# parametry drogi 
######
szer_drogi = excel['Wynik x'][1]
odległosc_od_ostatniej_warstwy = excel['Wynik x'][5]
liczba_warstw = excel['Wynik x'][18]
opisw1 = excel['Wynik x'][22]
opisw2 = excel['Wynik x'][23]
opisw3 = excel['Wynik x'][24]
opisw4 = excel['Wynik x'][25]
opisw5 = excel['Wynik x'][26]
dłg1 = excel['Wynik y'][22]
dłg2 = excel['Wynik y'][23]
dłg3 = excel['Wynik y'][24]
dłg4 = excel['Wynik y'][25]
dłg5 = excel['Wynik y'][26]
######
                                        #Punkty
##################################################################################################################
    #Budulcowe
a = APoint(excel['Wynik x'][2],excel['Wynik y'][2])
b = APoint(excel['Wynik x'][3],excel['Wynik y'][3])
c = APoint(excel['Wynik x'][4],excel['Wynik y'][4])
w1 = APoint(excel['Wynik x'][8],excel['Wynik y'][8])
z1 = APoint(excel['Wynik x'][9],excel['Wynik y'][9])
w2 = APoint(excel['Wynik x'][10],excel['Wynik y'][10])
z2 = APoint(excel['Wynik x'][11],excel['Wynik y'][11])
w3 = APoint(excel['Wynik x'][12],excel['Wynik y'][12])
z3 = APoint(excel['Wynik x'][13],excel['Wynik y'][13])
w4 = APoint(excel['Wynik x'][14],excel['Wynik y'][14])
z4 = APoint(excel['Wynik x'][15],excel['Wynik y'][15])
w5 = APoint(excel['Wynik x'][16],excel['Wynik y'][16])
z5 = APoint(excel['Wynik x'][17],excel['Wynik y'][17])
r1 = APoint(excel['Wynik x'][19],excel['Wynik y'][19]) 
r2 = APoint(excel['Wynik x'][20],excel['Wynik y'][20])
    #Opisowe
op0_1 = APoint(excel['Wynik x'][2]+8,excel['Wynik y'][2]+8)
op0_2 = APoint(excel['Wynik x'][2]+28,excel['Wynik y'][2]+8)
op1_1 = APoint(excel['Wynik x'][3]-8,excel['Wynik y'][3]+8)
op1_2 = APoint(excel['Wynik x'][3]-28,excel['Wynik y'][3]+8)
op2_1 = APoint(excel['Wynik x'][4]+8,excel['Wynik y'][4]+8)
op2_2 = APoint(excel['Wynik x'][4]+28,excel['Wynik y'][4]+8)
op3_1 = APoint(excel['Wynik x'][19]+8,excel['Wynik y'][19]-8)
op3_2 = APoint(excel['Wynik x'][19]+28,excel['Wynik y'][19]-8)
#########################################################################################################################



if Rodzaj_drogi != 'Glowna':

#pierwsza warstwa i opisy    
    l1 = acad.model.AddLine(a, b)
    l2 = acad.model.AddLine(b, c)
    l3 = acad.model.AddLine(c, r1)
    l4 = acad.model.AddLine(r1, r2)
    l1w1 = acad.model.AddLine(APoint(0,0-dłg1), w1)
    l2w1 = acad.model.AddLine(w1,z1)
                                #mirror
    l1.Mirror(APoint(0,0),APoint(0,-100))
    l2.Mirror(APoint(0,0),APoint(0,-100))
    l3.Mirror(APoint(0,0),APoint(0,-100))
    l4.Mirror(APoint(0,0),APoint(0,-100))
    l1w1.Mirror(APoint(0,0),APoint(0,-100))
    l2w1.Mirror(APoint(0,0),APoint(0,-100))
    #Opisy  
    opis_punkt_0_1 = acad.model.Addline(a,op0_1)
    opis_punkt_0_2 = acad.model.Addline(op0_1,op0_2)
    opis_punkt_1_1 = acad.model.Addline(b,op1_1)
    opis_punkt_1_2 = acad.model.Addline(op1_1,op1_2)
    opis_punkt_2_1 = acad.model.Addline(c,op2_1)
    opis_punkt_2_2 = acad.model.Addline(op2_1,op2_2)
    opis_punkt_3_1 = acad.model.Addline(r1,op3_1)
    opis_punkt_3_2 = acad.model.Addline(op3_1,op3_2)
    tekst_punktu_0 = acad.model.AddText(str(excel['Wynik y'][2]) + '.00', op0_1 ,6)
    tekst_punktu_1 = acad.model.AddText(str(round(excel['Wynik y'][3]/100,2)), op1_2 ,6)
    tekst_punktu_2 = acad.model.AddText(str(round(excel['Wynik y'][4]/100,2)), op2_1 ,6)
    tekst_punktu_3 = acad.model.AddText(str(round(excel['Wynik y'][19]/100,2)), op3_1 ,6)

                                #mirror
    
    opis_punkt_0_1.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_0_2.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_1_1.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_1_2.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_2_1.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_2_2.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_3_1.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_3_2.Mirror(APoint(0,0),APoint(0,-100))
    tekst_punktu_0.Mirror(APoint(0,0),APoint(0,-100))
    tekst_punktu_1.Mirror(APoint(0,0),APoint(0,-100))
    tekst_punktu_2.Mirror(APoint(0,0),APoint(0,-100))
    tekst_punktu_3.Mirror(APoint(0,0),APoint(0,-100))

                                #Strzałka nr 1
    #strzałka tworzenie
    strz_s =  APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,(excel['Wynik y'][2]+excel['Wynik y'][3])/2) #środek odcinka
    strz = l1.copy()
    strz_1 = acad.model.Addline(APoint(excel['Wynik x'][3],excel['Wynik y'][3]),APoint(excel['Wynik x'][3]-30,excel['Wynik y'][3])+10,) 
    strz_2 = strz_1.Mirror(a,b)
    strz_text1 = acad.model.AddText('5%', strz_s, 40 )
    #strzałka poruszenie

    strz.move(strz_s,APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+16))
    strz_1.move(strz_s,APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+16))
    strz_2.move(strz_s,APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+16))
    strz_text1.move(strz_s,APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+16))

    #strzałka skalowanie
    strz.ScaleEntity(strz_s , 0.19)
    strz_1.ScaleEntity(strz_s , 0.19)
    strz_2.ScaleEntity(strz_s , 0.19)
    strz_text1.ScaleEntity(APoint(((excel['Wynik x'][2]+excel['Wynik x'][3])/2)-10,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+1), 0.19)

                                                    #Strzałka nr 2
    strz_s =  APoint((excel['Wynik x'][3]+excel['Wynik x'][4])/2,(excel['Wynik y'][3]+excel['Wynik y'][4])/2) #środek odcinka
        #Kopiowanie
    strz2 =strz.copy() 
    strz_1_2 =strz_1.copy()
    strz_2_2 =strz_2.copy()
        #Poruszanie
    strz2.move(APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,(excel['Wynik y'][2]+excel['Wynik y'][3])/2),strz_s)
    strz_1_2.move(APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,(excel['Wynik y'][2]+excel['Wynik y'][3])/2),strz_s)
    strz_2_2.move(APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,(excel['Wynik y'][2]+excel['Wynik y'][3])/2),strz_s)
    strz_text2 = acad.model.AddText('6%',APoint(((excel['Wynik x'][3]+excel['Wynik x'][4])/2)-3,((excel['Wynik y'][3]+excel['Wynik y'][4])/2)+4) , 6 )
    
    #Mirror
    strz.Mirror(APoint(0,0),APoint(0,-100))
    strz_1.Mirror(APoint(0,0),APoint(0,-100))
    strz_2.Mirror(APoint(0,0),APoint(0,-100))
    strz_text1.Mirror(APoint(0,0),APoint(0,-100))
    
    strz2.Mirror(APoint(0,0),APoint(0,-100))
    strz_1_2.Mirror(APoint(0,0),APoint(0,-100))
    strz_2_2.Mirror(APoint(0,0),APoint(0,-100))
    strz_text2.Mirror(APoint(0,0),APoint(0,-100))

#Tworzenie warstw
    if liczba_warstw == 2:
        l1w2 = acad.model.AddLine(APoint(0,0-dłg1-dłg2), w2)
        l2w2 = acad.model.AddLine(w2,z2)
        opis_rowu_1 = acad.model.Addline(z2,APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
        opis_rowu_2 = acad.model.Addline(APoint(excel['Wynik x'][11],excel['Wynik y'][19]),r1)
        opis_rowu_3 = acad.model.Addline(APoint(excel['Wynik x'][11]-20,excel['Wynik y'][19]),APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
        Tekst_rowu = acad.model.AddText(str(odległosc_od_ostatniej_warstwy/100),APoint(excel['Wynik x'][11]-20,excel['Wynik y'][19]),5)
        
                                        #mirror

        l1w2.Mirror(APoint(0,0),APoint(0,-100))
        l2w2.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_1.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_2.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_3.Mirror(APoint(0,0),APoint(0,-100))
        Tekst_rowu.Mirror(APoint(0,0),APoint(0,-100))
        
    elif liczba_warstw == 3:
                                        #3 warstwa
        l1w2 = acad.model.AddLine(APoint(0,0-dłg1-dłg2), w2)
        l2w2 = acad.model.AddLine(w2,z2)
        
                                        #mirror

        l1w2.Mirror(APoint(0,0),APoint(0,-100))
        l2w2.Mirror(APoint(0,0),APoint(0,-100))
        
                                        #3 warstwa
        l1w3 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3), w3)
        l2w3 = acad.model.AddLine(w3,z3)
        opis_rowu_1 = acad.model.Addline(z3,APoint(excel['Wynik x'][13],excel['Wynik y'][19]))
        opis_rowu_2 = acad.model.Addline(APoint(excel['Wynik x'][13],excel['Wynik y'][19]),r1)
        opis_rowu_3 = acad.model.Addline(APoint(excel['Wynik x'][13]-20,excel['Wynik y'][19]),APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
        Tekst_rowu = acad.model.AddText(str(odległosc_od_ostatniej_warstwy/100),APoint(excel['Wynik x'][13]-20,excel['Wynik y'][19]),5)
                                            #mirror

        l1w3.Mirror(APoint(0,0),APoint(0,-100))
        l2w3.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_1.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_2.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_3.Mirror(APoint(0,0),APoint(0,-100))
        Tekst_rowu.Mirror(APoint(0,0),APoint(0,-100))
        
    elif liczba_warstw == 4:
                                        #3 warstwa
        l1w2 = acad.model.AddLine(APoint(0,0-dłg1-dłg2), w2)
        l2w2 = acad.model.AddLine(w2,z2)
        
                                        #mirror

        l1w2.Mirror(APoint(0,0),APoint(0,-100))
        l2w2.Mirror(APoint(0,0),APoint(0,-100))
        
                                        #3 warstwa
        l1w3 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3), w3)
        l2w3 = acad.model.AddLine(w3,z3)
                                            #mirror

        l1w3.Mirror(APoint(0,0),APoint(0,-100))
        l2w3.Mirror(APoint(0,0),APoint(0,-100))
        
                                        #4 warstwa
        l1w4 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3-dłg4), w4)
        l2w4 = acad.model.AddLine(w4,z4)
        opis_rowu_1 = acad.model.Addline(z4,APoint(excel['Wynik x'][15],excel['Wynik y'][19]))
        opis_rowu_2 = acad.model.Addline(APoint(excel['Wynik x'][15],excel['Wynik y'][19]),r1)
        opis_rowu_3 = acad.model.Addline(APoint(excel['Wynik x'][15]-20,excel['Wynik y'][19]),APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
        Tekst_rowu = acad.model.AddText(str(odległosc_od_ostatniej_warstwy/100),APoint(excel['Wynik x'][15]-20,excel['Wynik y'][19]),5)
                                            #mirror

        l1w4.Mirror(APoint(0,0),APoint(0,-100))
        l2w4.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_1.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_2.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_3.Mirror(APoint(0,0),APoint(0,-100))
        Tekst_rowu.Mirror(APoint(0,0),APoint(0,-100))
        
    elif liczba_warstw == 5: 

                                        #3 warstwa
        l1w2 = acad.model.AddLine(APoint(0,0-dłg1-dłg2), w2)
        l2w2 = acad.model.AddLine(w2,z2)
        
                                        #mirror

        l1w2.Mirror(APoint(0,0),APoint(0,-100))
        l2w2.Mirror(APoint(0,0),APoint(0,-100))
        
                                        #3 warstwa
        l1w3 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3), w3)
        l2w3 = acad.model.AddLine(w3,z3)
                                            #mirror

        l1w3.Mirror(APoint(0,0),APoint(0,-100))
        l2w3.Mirror(APoint(0,0),APoint(0,-100))
        
                                        #4 warstwa
        l1w4 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3-dłg4), w4)
        l2w4 = acad.model.AddLine(w4,z4)
                                            #mirror

        l1w4.Mirror(APoint(0,0),APoint(0,-100))
        l2w4.Mirror(APoint(0,0),APoint(0,-100))
                                            # 5 warstwa    
        l1w5 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3-dłg4-dłg5), w5)
        l2w5 = acad.model.AddLine(w5,z5)
        opis_rowu_1 = acad.model.Addline(z5,APoint(excel['Wynik x'][17],excel['Wynik y'][19]))
        opis_rowu_2 = acad.model.Addline(APoint(excel['Wynik x'][17],excel['Wynik y'][19]),r1)
        opis_rowu_3 = acad.model.Addline(APoint(excel['Wynik x'][17]-20,excel['Wynik y'][19]),APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
        Tekst_rowu = acad.model.AddText(str(odległosc_od_ostatniej_warstwy/100),APoint(excel['Wynik x'][17]-20,excel['Wynik y'][19]),5)
                                            #mirror

        l1w5.Mirror(APoint(0,0),APoint(0,-100))
        l2w5.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_1.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_2.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_3.Mirror(APoint(0,0),APoint(0,-100))
        Tekst_rowu.Mirror(APoint(0,0),APoint(0,-100))
        
    else:
        print('Nieprawidłowa liczba warstw (albo jedna warstwa)')
        opis_rowu_1 = acad.model.Addline(z1,APoint(excel['Wynik x'][9],excel['Wynik y'][19]))
        opis_rowu_2 = acad.model.Addline(APoint(excel['Wynik x'][9],excel['Wynik y'][19]),r1)
        opis_rowu_3 = acad.model.Addline(APoint(excel['Wynik x'][9]-20,excel['Wynik y'][19]),APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
        Tekst_rowu = acad.model.AddText(str(odległosc_od_ostatniej_warstwy/100),APoint(excel['Wynik x'][9]-20,excel['Wynik y'][19]),5)

#tu jest wariant koryta wtórnego
else:
        #pierwsza warstwa i opisy    
    l1 = acad.model.AddLine(a, b)
    l2 = acad.model.AddLine(b, c)
    l3 = acad.model.AddLine(c, r1)
    l4 = acad.model.AddLine(r1, r2)
   
                                #mirror
    l1.Mirror(APoint(0,0),APoint(0,-100))
    l2.Mirror(APoint(0,0),APoint(0,-100))
    l3.Mirror(APoint(0,0),APoint(0,-100))
    l4.Mirror(APoint(0,0),APoint(0,-100))
    
    #Opisy  
    opis_punkt_0_1 = acad.model.Addline(a,op0_1)
    opis_punkt_0_2 = acad.model.Addline(op0_1,op0_2)
    opis_punkt_1_1 = acad.model.Addline(b,op1_1)
    opis_punkt_1_2 = acad.model.Addline(op1_1,op1_2)
    opis_punkt_2_1 = acad.model.Addline(c,op2_1)
    opis_punkt_2_2 = acad.model.Addline(op2_1,op2_2)
    opis_punkt_3_1 = acad.model.Addline(r1,op3_1)
    opis_punkt_3_2 = acad.model.Addline(op3_1,op3_2)
    tekst_punktu_0 = acad.model.AddText(str(excel['Wynik y'][2]) + '.00', op0_1 ,6)
    tekst_punktu_1 = acad.model.AddText(str(round(excel['Wynik y'][3]/100,2)), op1_2 ,6)
    tekst_punktu_2 = acad.model.AddText(str(round(excel['Wynik y'][4]/100,2)), op2_1 ,6)
    tekst_punktu_3 = acad.model.AddText(str(round(excel['Wynik y'][19]/100,2)), op3_1 ,6)

                                #mirror
    
    opis_punkt_0_1.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_0_2.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_1_1.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_1_2.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_2_1.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_2_2.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_3_1.Mirror(APoint(0,0),APoint(0,-100))
    opis_punkt_3_2.Mirror(APoint(0,0),APoint(0,-100))
    tekst_punktu_0.Mirror(APoint(0,0),APoint(0,-100))
    tekst_punktu_1.Mirror(APoint(0,0),APoint(0,-100))
    tekst_punktu_2.Mirror(APoint(0,0),APoint(0,-100))
    tekst_punktu_3.Mirror(APoint(0,0),APoint(0,-100))

                                #Strzałka nr 1
    #strzałka tworzenie
    strz_s =  APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,(excel['Wynik y'][2]+excel['Wynik y'][3])/2) #środek odcinka
    strz = l1.copy()
    strz_1 = acad.model.Addline(APoint(excel['Wynik x'][3],excel['Wynik y'][3]),APoint(excel['Wynik x'][3]-30,excel['Wynik y'][3])+10,) 
    strz_2 = strz_1.Mirror(a,b)
    strz_text1 = acad.model.AddText('5%', strz_s, 40 )
    #strzałka poruszenie

    strz.move(strz_s,APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+16))
    strz_1.move(strz_s,APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+16))
    strz_2.move(strz_s,APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+16))
    strz_text1.move(strz_s,APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+16))

    #strzałka skalowanie
    strz.ScaleEntity(strz_s , 0.19)
    strz_1.ScaleEntity(strz_s , 0.19)
    strz_2.ScaleEntity(strz_s , 0.19)
    strz_text1.ScaleEntity(APoint(((excel['Wynik x'][2]+excel['Wynik x'][3])/2)-10,((excel['Wynik y'][2]+excel['Wynik y'][3])/2)+1), 0.19)

                                                    #Strzałka nr 2
    strz_s =  APoint((excel['Wynik x'][3]+excel['Wynik x'][4])/2,(excel['Wynik y'][3]+excel['Wynik y'][4])/2) #środek odcinka
        #Kopiowanie
    strz2 =strz.copy() 
    strz_1_2 =strz_1.copy()
    strz_2_2 =strz_2.copy()
        #Poruszanie
    strz2.move(APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,(excel['Wynik y'][2]+excel['Wynik y'][3])/2),strz_s)
    strz_1_2.move(APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,(excel['Wynik y'][2]+excel['Wynik y'][3])/2),strz_s)
    strz_2_2.move(APoint((excel['Wynik x'][2]+excel['Wynik x'][3])/2,(excel['Wynik y'][2]+excel['Wynik y'][3])/2),strz_s)
    strz_text2 = acad.model.AddText('6%',APoint(((excel['Wynik x'][3]+excel['Wynik x'][4])/2)-3,((excel['Wynik y'][3]+excel['Wynik y'][4])/2)+4) , 6 )
    
    #Mirror
    strz.Mirror(APoint(0,0),APoint(0,-100))
    strz_1.Mirror(APoint(0,0),APoint(0,-100))
    strz_2.Mirror(APoint(0,0),APoint(0,-100))
    strz_text1.Mirror(APoint(0,0),APoint(0,-100))
    
    strz2.Mirror(APoint(0,0),APoint(0,-100))
    strz_1_2.Mirror(APoint(0,0),APoint(0,-100))
    strz_2_2.Mirror(APoint(0,0),APoint(0,-100))
    strz_text2.Mirror(APoint(0,0),APoint(0,-100))

    if Nawierzchnia == 'tłuczniowa':
        l1w1 = acad.model.AddLine(APoint(0,0-dłg1), w1)
        l2w1 = acad.model.AddLine(b,w1)
        l1w1.Mirror(APoint(0,0),APoint(0,-100))
        l2w1.Mirror(APoint(0,0),APoint(0,-100))
        
        if liczba_warstw == 2:
            #druga warstwa
            l1w2 = acad.model.AddLine(APoint(0,0-dłg1-dłg2), w2)
            l2w2 = acad.model.AddLine(w1,w2)
            opis_rowu_1 = acad.model.Addline(z2,APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
            opis_rowu_2 = acad.model.Addline(APoint(excel['Wynik x'][11],excel['Wynik y'][19]),r1)
            opis_rowu_3 = acad.model.Addline(APoint(excel['Wynik x'][11]-20,excel['Wynik y'][19]),APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
            Tekst_rowu = acad.model.AddText(str(odległosc_od_ostatniej_warstwy/100),APoint(excel['Wynik x'][11]-20,excel['Wynik y'][19]),5)
            
                                            #mirror

            l1w2.Mirror(APoint(0,0),APoint(0,-100))
            l2w2.Mirror(APoint(0,0),APoint(0,-100))
            opis_rowu_1.Mirror(APoint(0,0),APoint(0,-100))
            opis_rowu_2.Mirror(APoint(0,0),APoint(0,-100))
            opis_rowu_3.Mirror(APoint(0,0),APoint(0,-100))
            Tekst_rowu.Mirror(APoint(0,0),APoint(0,-100))
            
        else:
            #druga warstwa (korytowy system)
            l1w2 = acad.model.AddLine(APoint(0,0-dłg1-dłg2), w2)
            l2w2 = acad.model.AddLine(w1,w2)
                                            
                                                #mirror
            l1w2.Mirror(APoint(0,0),APoint(0,-100))
            l2w2.Mirror(APoint(0,0),APoint(0,-100))
            
            #warstwa odcinająco-odsączająca
            l1w3 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3), w3)
            l2w3 = acad.model.AddLine(w3,z3)
            opis_rowu_1 = acad.model.Addline(z3,APoint(excel['Wynik x'][13],excel['Wynik y'][19]))
            opis_rowu_2 = acad.model.Addline(APoint(excel['Wynik x'][13],excel['Wynik y'][19]),r1)
            opis_rowu_3 = acad.model.Addline(APoint(excel['Wynik x'][13]-20,excel['Wynik y'][19]),APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
            Tekst_rowu = acad.model.AddText(str(odległosc_od_ostatniej_warstwy/100),APoint(excel['Wynik x'][13]-20,excel['Wynik y'][19]),5)
                                                #mirror

            l1w3.Mirror(APoint(0,0),APoint(0,-100))
            l2w3.Mirror(APoint(0,0),APoint(0,-100))
            opis_rowu_1.Mirror(APoint(0,0),APoint(0,-100))
            opis_rowu_2.Mirror(APoint(0,0),APoint(0,-100))
            opis_rowu_3.Mirror(APoint(0,0),APoint(0,-100))
            Tekst_rowu.Mirror(APoint(0,0),APoint(0,-100))

    else:
        
        l1w1 = acad.model.AddLine(APoint(0,0-dłg1), w1)
        l2w1 = acad.model.AddLine(b,w1)
        l1w1.Mirror(APoint(0,0),APoint(0,-100))
        l2w1.Mirror(APoint(0,0),APoint(0,-100))
                
            #druga warstwa
        l1w2 = acad.model.AddLine(APoint(0,0-dłg1-dłg2), w2)
        l2w2 = acad.model.AddLine(w1,w2)
            #trzecia warstwa
        l1w3 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3), w3)
        l2w3 = acad.model.AddLine(w2,w3)
            #czwarta warstwa
        l1w4 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3-dłg4), w4)
        l2w4 = acad.model.AddLine(w3,w4)

        #warstwa odsączająco-odcinająca 

        l1w5 = acad.model.AddLine(APoint(0,0-dłg1-dłg2-dłg3-dłg4-dłg5), w5)
        l2w5 = acad.model.AddLine(w5,z5)
        opis_rowu_1 = acad.model.Addline(z5,APoint(excel['Wynik x'][17],excel['Wynik y'][19]))
        opis_rowu_2 = acad.model.Addline(APoint(excel['Wynik x'][17],excel['Wynik y'][19]),r1)
        opis_rowu_3 = acad.model.Addline(APoint(excel['Wynik x'][17]-20,excel['Wynik y'][19]),APoint(excel['Wynik x'][11],excel['Wynik y'][19]))
        Tekst_rowu = acad.model.AddText(str(odległosc_od_ostatniej_warstwy/100),APoint(excel['Wynik x'][17]-20,excel['Wynik y'][19]),5)
                                        #mirror

        l1w2.Mirror(APoint(0,0),APoint(0,-100))
        l2w2.Mirror(APoint(0,0),APoint(0,-100))
        l1w3.Mirror(APoint(0,0),APoint(0,-100))
        l2w3.Mirror(APoint(0,0),APoint(0,-100))
        l1w4.Mirror(APoint(0,0),APoint(0,-100))
        l2w4.Mirror(APoint(0,0),APoint(0,-100))
        l1w5.Mirror(APoint(0,0),APoint(0,-100))
        l2w5.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_1.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_2.Mirror(APoint(0,0),APoint(0,-100))
        opis_rowu_3.Mirror(APoint(0,0),APoint(0,-100))
        Tekst_rowu.Mirror(APoint(0,0),APoint(0,-100))
            
###########################################################################################################################################################################
#                                                                                           Wymiary drogi

#               primy

b = APoint(excel['Wynik x'][3],excel['Wynik y'][3])
bprim = APoint(-excel['Wynik x'][3],excel['Wynik y'][3])
c = APoint(excel['Wynik x'][4],excel['Wynik y'][4])
cprim = APoint(-excel['Wynik x'][4],excel['Wynik y'][4])

#               góra

bgora = APoint(excel['Wynik x'][3],excel['Wynik y'][3]+60)
bprimgora = APoint(-excel['Wynik x'][3],excel['Wynik y'][3]+60)
cgora = APoint(excel['Wynik x'][4],excel['Wynik y'][3]+60)
cprimgora = APoint(-excel['Wynik x'][4],excel['Wynik y'][3]+60)

#               Przedłużenie

bprzed = APoint(excel['Wynik x'][3],excel['Wynik y'][3]+65)
bprimprzed = APoint(-excel['Wynik x'][3],excel['Wynik y'][3]+65)
cprzed = APoint(excel['Wynik x'][4],excel['Wynik y'][3]+65)
cprimprzed = APoint(-excel['Wynik x'][4],excel['Wynik y'][3]+65)

#               pionowe linie
wymiar1 = acad.model.AddLine(b,bgora)
wymiar2 = acad.model.AddLine(bprim,bprimgora)
wymiar3 = acad.model.AddLine(c,cgora)
wymiar4 = acad.model.AddLine(cprim,cprimgora)
wymiar5 = acad.model.AddLine(bprimgora,bprimprzed)
wymiar6 = acad.model.AddLine(cprimgora,cprimprzed)
wymiar7 = acad.model.AddLine(bgora,bprzed)
wymiar8 = acad.model.AddLine(cgora,cprzed)
#               poziome linie
wymiar9 = acad.model.AddLine(cprimgora,bprimgora)
wymiar10 = acad.model.AddLine(bprimgora,bgora)  
wymiar11 = acad.model.AddLine(bgora,cgora)

if szer_drogi == 350:
    tekst_szer = acad.model.AddText('3.50', APoint(-7,55), 6)    
    tekst_pobocza1 = acad.model.AddText('0.50', APoint(192,55), 6)
    tekst_pobocza2 = acad.model.AddText('0.50', APoint(-207,55), 6)
else:
    tekst_szer = acad.model.AddText('3.50', APoint(-7,55), 6)    
    tekst_pobocza1 = acad.model.AddText('0.50', APoint(168,55), 6)
    tekst_pobocza2 = acad.model.AddText('0.50', APoint(-183,55), 6)


                                #Tytuł i skala
if Rodzaj_drogi == 'Glowna': 
    if Grunt == 'sypkie':
        
        tytul = acad.model.AddText('Przekrój normalny na prostej - droga główna (grunt sypki)',APoint(-178,174),10)
    
    elif Grunt == 'mało spoiste': 
         
         tytul = acad.model.AddText('Przekrój normalny na prostej - droga główna (grunt mało spoisty)',APoint(-178,174),10)
    
    else:
         
         tytul = acad.model.AddText('Przekrój normalny na prostej - droga główna (grunt spoisty)',APoint(-178,174),10)
elif Rodzaj_drogi == 'Boczna':
    if Grunt == 'sypkie':
        
        tytul = acad.model.AddText('Przekrój normalny na prostej - droga boczna (grunt sypki)',APoint(-178,174),10)
    
    elif Grunt == 'mało spoiste': 
         
         tytul = acad.model.AddText('Przekrój normalny na prostej - droga boczna (grunt mało spoisty)',APoint(-178,174),10)
    
    else:
         
         tytul = acad.model.AddText('Przekrój normalny na prostej - droga boczna (grunt spoisty)',APoint(-178,174),10)
else:
    if Grunt == 'sypkie':
        
        tytul = acad.model.AddText('Przekrój normalny na prostej - droga łącznikowa (grunt sypki)',APoint(-178,174),10)
    
    elif Grunt == 'mało spoiste': 
         
         tytul = acad.model.AddText('Przekrój normalny na prostej - droga łącznikowa (grunt mało spoisty)',APoint(-178,174),10)
    
    else:
         
         tytul = acad.model.AddText('Przekrój normalny na prostej - droga łącznikowa (grunt spoisty)',APoint(-178,174),10)


skala = acad.model.AddText('Skala 1:50',APoint(228,145),10)

#Opis warstw DOPRACOWAĆ
w1_1 = APoint(0,0-(dłg1/2))
w1_2 = APoint(0+100,0-(dłg1/2)-70)
w1_3 = APoint(0+250,0-(dłg1/2)-70)

w2_1 = APoint(0,0-dłg1-(dłg2/2))
w2_2 = APoint(0+100,0-dłg1-(dłg2/2)-70)
w2_3 = APoint(0+250,0-dłg1-(dłg2/2)-70)

w3_1 = APoint(0,0-dłg1-dłg2-(dłg3/2))
w3_2 = APoint(0+100,0-dłg1-dłg2-(dłg3/2)-70)
w3_3 = APoint(0+250,0-dłg1-dłg2-(dłg3/2)-70)

w4_1 = APoint(0,0-dłg1-dłg2-dłg3-(dłg4/2))
w4_2 = APoint(0+100,0-dłg1-dłg2-dłg3-(dłg4/2)-70)
w4_3 = APoint(0+250,0-dłg1-dłg2-dłg3-(dłg4/2)-70)

w5_1 = APoint(0,0-dłg1-dłg2-dłg3-dłg4-(dłg5/2))
w5_2 = APoint(0+100,0-dłg1-dłg2-dłg3-dłg4-(dłg5/2)-70)
w5_3 = APoint(0+250,0-dłg1-dłg2-dłg3-dłg4-(dłg5/2)-70)
                          
if liczba_warstw == 1:
    acad.model.AddLine(w1_1,w1_2)
    acad.model.AddLine(w1_2,w1_3)
    acad.model.AddText(opisw1 ,w1_2, 4)
elif liczba_warstw == 2:
    acad.model.AddLine(w1_1,w1_2)
    acad.model.AddLine(w1_2,w1_3)
    acad.model.AddText(opisw1 ,w1_2, 4)
    acad.model.AddLine(w2_1,w2_2)
    acad.model.AddLine(w2_2,w2_3)
    acad.model.AddText(opisw2 ,w2_2, 4)
elif liczba_warstw == 3:
    acad.model.AddLine(w1_1,w1_2)
    acad.model.AddLine(w1_2,w1_3)
    acad.model.AddText(opisw1 ,w1_2, 4)
    acad.model.AddLine(w2_1,w2_2)
    acad.model.AddLine(w2_2,w2_3)
    acad.model.AddText(opisw2 ,w2_2, 4)
    acad.model.AddLine(w3_1,w3_2)
    acad.model.AddLine(w3_2,w3_3)
    acad.model.AddText(opisw3 ,w3_2, 4)
elif liczba_warstw == 4:
    acad.model.AddLine(w1_1,w1_2)
    acad.model.AddLine(w1_2,w1_3)
    acad.model.AddText(opisw1 ,w1_2, 4)
    acad.model.AddLine(w2_1,w2_2)
    acad.model.AddLine(w2_2,w2_3)
    acad.model.AddText(opisw2 ,w2_2, 4)
    acad.model.AddLine(w3_1,w3_2)
    acad.model.AddLine(w3_2,w3_3)
    acad.model.AddText(opisw3 ,w3_2, 4)
    acad.model.AddLine(w4_1,w4_2)
    acad.model.AddLine(w4_2,w4_3)
    acad.model.AddText(opisw4 ,w4_2, 4)
else:
    acad.model.AddLine(w1_1,w1_2)
    acad.model.AddLine(w1_2,w1_3)
    acad.model.AddText(opisw1 ,w1_2, 4)
    acad.model.AddLine(w2_1,w2_2)
    acad.model.AddLine(w2_2,w2_3)
    acad.model.AddText(opisw2 ,w2_2, 4)
    acad.model.AddLine(w3_1,w3_2)
    acad.model.AddLine(w3_2,w3_3)
    acad.model.AddText(opisw3 ,w3_2, 4)
    acad.model.AddLine(w4_1,w4_2)
    acad.model.AddLine(w4_2,w4_3)
    acad.model.AddText(opisw4 ,w4_2, 4)
    acad.model.AddLine(w5_1,w5_2)
    acad.model.AddLine(w5_2,w5_3)
    acad.model.AddText(opisw5 ,w5_2, 4)



