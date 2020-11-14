# Battleground Data Visualization & Statistics

## Overview

This is a data visualization and statistics tool for Hearthstone Battleground. 
It provides statistics and visualization for data on an Excel file. In this file, data correspond to your manually entered results over time.
Two operations are available for the moment : 
*   Web Scraping average ranking of each heroe on [HS Replay](https://hsreplay.net/battlegrounds/heroes/).
*   Analyse compositions of minions for each heroe on the Excel file with a visual rendering.


### Web Scraping

<img src = "images/hs_replay.png?raw=true" />

Extracting and placing Heroe's average ranking in the Excel file is automatically managed by webscrapping. 

### Statistics 

<img src="images/Edwin_result.png?raw=true" />

This application provides an analysis of compositions and results for each available heroe. 
The visual rendering is done with Matplotlib library.
Personnal best comps and hero's efficiency are rendered. 

<img src="images/general_stats.png?raw=true" />


## Requirements 

*   Python 3.7 minimum
*   Excel template file will be avaible soon

Needed libraries in Python are listed in requirements.txt. Run the following command.

```
pip install -r requirements.txt 
```


