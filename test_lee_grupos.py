# -*- coding: utf-8 -*-
"""
Created on Fri Jul 30 17:56:26 2021

@author: Ruben
"""
import pandas as pd
from pathlib import Path

lista_subgrupos = pd.read_excel('lista_subgrupos.xlsx', index_col='Email').sort_index()

