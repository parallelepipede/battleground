#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jun 13 22:49:41 2020

@author: Paul Zanoncelli
"""


import xlwings as xw
from selenium import webdriver
import time
from bs4 import BeautifulSoup

wb = xw.Book('stat.xlsx') 
ws = wb.sheets[0]

number_heros = 44            # heroes on site
number_available_heros = 53

def webscrape():
    heros = dict()
    for i in range(17,number_available_heros * + 19 + 1, 9 ):
        a = ws.range('A' + str(i)).value
        if a is not None : 
            heros[a] = i
        
    stats_global = heros
    url = 'https://hsreplay.net/battlegrounds/heroes/'
    driver = webdriver.Safari()
    driver.get(url)
    stats_global = dict()

    def fill_stats_global(delay):
        global stats_global
        time.sleep(delay)
        html = driver.page_source
        soup = BeautifulSoup(html,'lxml')
        a = soup.prettify()
        b = a.split("bgs-table-cell__hero-right")
        for i in range(1,len(b)):
            stats_global[b[i].split("span")[1][2:-2].strip()] = b[i].split('<div class="bgs-data-placement')[1].split('>')[2].split('<')[0].strip()
        print()
        print(stats_global,len(b))
    
    fill_stats_global(12)
    driver.execute_script("window.scrollTo(document.body.scrollHeight/3,2*document.body.scrollHeight/3);")
    fill_stats_global(3)
    driver.execute_script("window.scrollTo(2*document.body.scrollHeight/3,document.body.scrollHeight);")
    fill_stats_global(3)
    print(str(len(stats_global)) + " heroes are downloaded")
    def update():
        for hero in heros : 
            for items in stats_global : 
                if hero.upper() in items.upper():
                    ws.range('E'+str(heros[hero] + 8)).value = "Statistiques globales"
                    ws.range('D'+str(heros[hero] + 8)).value = stats_global[items]
        
        if len(stats_global) == number_heros: 
            print("Every heroe is correctly downloaded")
        else:
            print("Missing heroes : ")
            for hero in heros : 
                here = False
                for items in stats_global : 
                    if hero.upper() in items.upper():
                        here = True
                        break
                if not here : 
                    print(hero)
        ws.range('P406').value = "Statistiques Globales : Héros"
        ws.range('Q406').value = "Rang Moyen"
        for i,items in enumerate(stats_global):
            ws.range('P' + str(500 + i)).value = str(i+1) + ' ' + items
            ws.range('Q' + str(500 + i)).value = stats_global[items] 
            
    update()
    driver.quit()
    
    