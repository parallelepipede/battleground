#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jun 13 22:49:41 2020

@author: paul
"""

import xlwings as xw
import matplotlib.pyplot as plt
import time

wb = xw.Book('stat.xlsx') 
ws = wb.sheets[0]

# TODO threading

number_heros = 44            # heroes on site
number_available_heros = 53  #heroes in file up to Zephrys

def best_comp():
    comp =   ["dragon","beast","mech","demon","pirate","menagerie", "divine","murloc","elem"]
    letter = [chr(i) for i in range(65,91)]
    size = len(letter)
    global_dict = {}
    for composition in comp : 
        global_dict[composition] = [0,0,[0]*8]
        
    for i in range(19,number_available_heros * 9 + 19 + 1 ,9): #389 for 41 heros
        j = 2
        hero_comp = {}
        for composition in comp : 
            hero_comp[composition] = [0,0]
        index = letter[j]
        while ws.range(index + str(i)).value : 
            compo = ws.range(index+ str(i)).value
            for composition in comp : 
                if composition in compo : 
                    hero_comp[composition][0] += 1
                    hero_comp[composition][1] += int(ws.range(index + str(i-1)).value)
                    global_dict[composition][2][int(ws.range(index + str(i-1)).value) - 1] += 1
                    break
            j += 1
            if j >= size : 
                index =  letter[j//size - 1] + letter[j % size ]
            else:
                index = letter[j]
        print(hero_comp,' ', ws.range('A' + str(i-2)).value)
        hero_comp = dict(filter(lambda item : item[1][0] > 0,hero_comp.items()))
        for composition in hero_comp : 
            global_dict[composition][0] += hero_comp[composition][0]
            global_dict[composition][1] += hero_comp[composition][1]
            hero_comp[composition][1] /= hero_comp[composition][0]
        for label,items in enumerate(dict(sorted(hero_comp.items(),key = lambda item : (item[1][0],-item[1][1]),reverse = True))):
            game = ' game' if hero_comp[items][0] == 1 else ' games'
            ws.range(letter[label + 6] + str(i + 5)).value = items + ' :  ' + str(hero_comp[items][0]) + game + ', ' + '%.2f' % hero_comp[items][1] + ' average rank' 
        for label,items in enumerate(dict(sorted(hero_comp.items(),key = lambda item : item[1][1]))):
            game = ' game' if hero_comp[items][0] == 1 else ' games'
            ws.range(letter[label + 6] + str(i + 6)).value = items + ' :  '  + '%.2f' % hero_comp[items][1] + ' average rank, ' +  str(hero_comp[items][0]) + game
        
    for i,composition in enumerate(global_dict):
        global_dict[composition][1] /= global_dict[composition][0]
        fig = plt.figure()
        plt.bar(range(1,9),global_dict[composition][2],align = 'center', alpha = .7,color = 'b',width = .5)
        plt.title(composition[0].upper() + composition[1:],fontsize = 'xx-large')
        plt.grid(axis = 'y')
        if ws.pictures[composition] : 
            ws.pictures[composition].delete()
        ws.pictures.add(fig,name = composition, update = False,left = ws.range(letter[i + 7] + '6').left,top = ws.range(letter[i+ 7] + '6').top,width = ws.range(letter[i + 7] + '6').width, height = 150)
        
    for label,items in enumerate(dict(sorted(global_dict.items(),key = lambda item : (item[1][0],-item[1][1]),reverse = True))):
        game = ' game' if global_dict[items][0] == 1 else ' games' 
        ws.range('F' + str(7 + label)).value = items + " :  " +str(global_dict[items][0]) + game + ', %.2f' % global_dict[items][1] + ' average rank'
        
    for label,items in enumerate(dict(sorted(global_dict.items(),key = lambda item : item[1][1]))):
        game = ' game' if global_dict[items][0] == 1 else ' games'
        ws.range('G' + str(7 + label)).value = items + " :  " +'%.2f' % global_dict[items][1] + ' average rank, ' + str(global_dict[items][0]) + game
tim = time.time()
best_comp()

print('Computing Time : ', time.time() - tim) 


