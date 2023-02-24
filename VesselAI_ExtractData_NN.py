#!/usr/bin/python
# -*- coding: <<encoding>> -*-
#-------------------------------------------------------------------------------
#   <<project>>
# 
#-------------------------------------------------------------------------------


# The program tries to find the relevant data pages (containing e.g. fuel consumption info) in the *PDFs listed in 'files_to_process'.
# When a relevant page is found, it is saved as a *.png and a *.html.
# All data lines in the relevant pages are printed to a file, from where the Wx UI extracts the data to show it.


# import wxversion
# wxversion.select("2.8")
import wx, wx.html
import sys
import webbrowser
import fitz
import os
import html
import unicodedata
import string
import json
import random

import nltk
from nltk.corpus import stopwords
import string
from string import punctuation
from os import listdir
from collections import Counter
from keras.preprocessing.text import Tokenizer
from numpy import array
from keras.models import Sequential
from keras.models import load_model
from keras.layers import Dense
from keras.layers import Dropout
from pandas import DataFrame
from matplotlib import pyplot
# import pickle # Tested vocabulary and tokenizer saving and reloading, but these pieces of code were commented away...

import os
os.environ["CUDA_VISIBLE_DEVICES"] = "-1" # disable GPU



aboutText = """<p>Sorry, there is no information about this program. It is
running on version %(wxpy)s of <b>wxPython</b> and %(python)s of <b>Python</b>.
See <a href="http://wiki.wxpython.org">wxPython Wiki</a></p>""" 

resultsFolder1 = "NN_Docs_Pages"


files_to_process = [
    "MAN/man-175d-imo-tier-ii-imo-tier-iii-marine.pdf",
    "MAN/man-32-40-imo-tier-ii-marine.pdf",
    "MAN/man-32-40-imo-tier-iii-marine.pdf",
    "MAN/man-32-44cr-imo-tier-ii-marine.pdf",
    "MAN/man-32-44cr-imo-tier-iii-marine.pdf",
    "MAN/man-48-60cr-imo-tier-ii-marine.pdf",
    "MAN/man-48-60cr-imo-tier-iii-marine.pdf",
    "MAN/man-51-60df-imo-tier-ii-imo-tier-iii-marine.pdf",
    "MAN/man-51-60df-imo-tier-iii-marine.pdf",
    "MAN/man-es_l3_factsheets_4p_man_175d_genset_n5_201119_low_new.pdf",
    "MAN/man-l32-44-genset-imo-tier-ii-marine.pdf",
    "MAN/man-l32-44-genset-imo-tier-iii-marine.pdf",
    "MAN/man-l35-44df-imo-tier-ii-marine.pdf",
    "MAN/man-l35-44df-imo-tier-iii-marine.pdf",
    "MAN/man-v28-33d-stc-imo-tier-ii-marine.pdf",
    "MAN/man-v28-33d-stc-imo-tier-iii-marine.pdf",
    "MAN/PG_M-III_L2131.pdf",
    "MAN/PG_M-II_L1624.pdf",
    "MAN/PG_P-III_L2131.pdf",
    "MAN/PG_P-II_L2131.pdf",
    "MAN/PG_P-I_L2131.pdf",
    
    "Wartsila/brochure-o-e-w31.pdf",
    "Wartsila/brochure-o-e-wartsila-dual-fual-low-speed.pdf",
    "Wartsila/leaflet-w14-2018.pdf",
    "Wartsila/marine-power-catalogue.pdf",
    "Wartsila/product-guide-o-e-w26.pdf",
    "Wartsila/product-guide-o-e-w31df.pdf",
    "Wartsila/product-guide-o-e-w32.pdf",
    "Wartsila/product-guide-o-e-w34df.pdf",
    "Wartsila/product-guide-o-e-w46df.pdf",
    "Wartsila/product-guide-o-e-w46f.pdf",
    "Wartsila/wartsila-14-product-guide.pdf",
    "Wartsila/wartsila-31sg-product-guide.pdf",
    "Wartsila/wartsila-46df.pdf",
    "Wartsila/wartsila-o-e-ls-x92.pdf",
    
    "Hyundai_HHI/2016_HHI_en.pdf",
    "Hyundai_HHI/2017_HHI_en.pdf",
    "Hyundai_HHI/2018_HHI_en.pdf",
    "Hyundai_HHI/2019_HHI_en.pdf",
    "Hyundai_HHI/2020_KSOE_IR_EN.pdf",
    "Hyundai_HHI/2021_KSOE_IR_EN.pdf",
    "Hyundai_HHI/HHI_EMD_brochure2017_1.pdf",
    "Hyundai_HHI/HHI_EMD_brochure2017_2.pdf",
    "Hyundai_HHI/HHI_EMD_brochure2017_3.pdf",
    "Hyundai_HHI/HHI_EMD_brochure2019_1.pdf",
    "Hyundai_HHI/speclal_NSD2020.pdf",
    
    "WinGD/16-06-pres-kit_x92.pdf",
    "WinGD/cimac2016_120_kyrtatos_paper_thedevelopmentofthemodern2strokemarinedieselengine.pdf",
    "WinGD/cimac2016_173_brueckl_paper_virtualdesign-and-simulation-in2strokemarineenginedevelopment.pdf",
    "WinGD/cimac2016_233_ott_paper_x-df.pdf",
    "WinGD/dominikschneiter_tieriii-programme.pdf",
    "WinGD/Fuel-Flexible-Injection-System-How-to-Handle-a-Fuel-Specturm-from-Diesel_CIMAC2019_paper_404_Andreas-Schmid.pdf",
    "WinGD/marcspahni_generation-x-engines.pdf",
    "WinGD/MIM_WinGD-RT-flex48T-D.pdf",
    "WinGD/MIM_WinGD-X72.pdf",
    "WinGD/MM_WinGD-RT-flex58T-D.pdf",
    "WinGD/MM_WinGD-X35-B_2021-09.pdf",
    "WinGD/MM_WinGD-X52.pdf",
    "WinGD/MM_WinGD-X72.pdf",
    "WinGD/MM_WinGD-X72DF.pdf",
    "WinGD/MM_WinGD-X82-B.pdf",
    "WinGD/Motorship-May-2018-VP-R-D-Leader-Brief.pdf",
    "WinGD/OM_WinGD-X62_2021-09.pdf",
    "WinGD/OM_WinGD-X72DF_2021-09.pdf",
    "WinGD/OM_WinGD-X82-B.pdf",
    "WinGD/OM_WinGD_RT-flex50DF.pdf",
    "WinGD/WinGD-12X92DF-Development-of-the-Most-Powerful-Otto-Engine-Ever_CIMAC2019_paper_425_Dominik-Schneiter.pdf",
    "WinGD/wingd-paper_engine_selection_for_very_large_container_vessels_201609.pdf",
    "WinGD/WinGD-WiDE-Brochure.pdf",
    "WinGD/WinGD_Engine-Booklet_2021.pdf",
    "WinGD/wingd_moving-inlet-ports-concept-for-optimization-of-2-stroke-uni-flow-engines_patrick-rebecchi.pdf",
    "WinGD/X-DF-Engines-by-WinGD.pdf",
    "WinGD/X-DF-FAQ.pdf"
    ]

# Relevant pages: Doc number, page number. By human expert, 690 pages
ground_truth_doc_pages = [
    # "MAN/man-175d-imo-tier-ii-imo-tier-iii-marine.pdf",
    [0, 79],
    [0, 80],
    [0, 81],
    [0, 83],
    [0, 84],
    [0, 85],
    [0, 86],
    [0, 87],
    [0, 88],
    [0, 90],
    [0, 91],
    [0, 92],
    [0, 94],
    [0, 95],
    [0, 96],
    [0, 98],
    [0, 99],
    [0, 100],
    [0, 102],
    [0, 103],
    [0, 104],
    [0, 106],
    [0, 107],
    [0, 108],
    [0, 110],
    [0, 111],
    [0, 112],
    [0, 114],
    [0, 115],
    [0, 116],
    [0, 118],
    [0, 119],
    [0, 120],
    [0, 122],
    [0, 123],
    [0, 124],
    [0, 126],
    [0, 127],
    [0, 128],
    [0, 130],
    [0, 131],
    [0, 132],
    [0, 134],
    [0, 135],
    [0, 136],
    [0, 138],
    [0, 139],
    [0, 140],
    [0, 142],
    [0, 143],
    [0, 144],
    [0, 146],
    [0, 147],
    [0, 148],
    [0, 150],
    [0, 151],
    [0, 152],
    [0, 154],
    [0, 155],
    [0, 156],
    [0, 158],
    [0, 159],
    [0, 160],
    [0, 162],
    [0, 163],
    [0, 164],
    [0, 166],
    [0, 167],
    [0, 168],
    [0, 170],
    [0, 171],
    [0, 172],
    [0, 174],
    [0, 175],
    [0, 176],
    [0, 178],
    [0, 179],
    [0, 180],
    [0, 182],
    [0, 183],
    [0, 184],
    [0, 185],
    [0, 186],
    [0, 187],
    [0, 188],
    [0, 189],
    [0, 190],
    [0, 191],
    [0, 192],
    [0, 193],
    [0, 194],
    [0, 195],
    [0, 196],
    [0, 197],
    [0, 198],
    [0, 199],
    [0, 200],
    [0, 202],
    [0, 203],
    [0, 204],
    [0, 205],
    [0, 206],
    [0, 207],
    [0, 208],
    [0, 209],
    [0, 210],
    [0, 212],
    [0, 213],
    [0, 214],
    [0, 216],
    [0, 217],
    [0, 218],
    [0, 220],

    # "MAN/man-32-40-imo-tier-ii-marine.pdf",
    [1, 78],
    [1, 79],
    [1, 80],
    [1, 83],
    [1, 87],
    [1, 88],
    [1, 89],
    [1, 90],
    [1, 91],
    [1, 92],
    [1, 93],
    [1, 94],
    [1, 95],
    [1, 96],
    [1, 97],
    [1, 98],
    [1, 99],
    [1, 100],
    [1, 101],
    [1, 102],
    [1, 103],
    [1, 104],
    [1, 105],
    [1, 106],
    [1, 107],
    [1, 108],
    [1, 109],
    [1, 110],
    [1, 111],
    [1, 112],
    [1, 113],
    [1, 114],
    [1, 115],
    [1, 116],
    [1, 117],
    [1, 118],
    [1, 119],
    [1, 120],
    [1, 121],
    [1, 122],
    [1, 123],
    [1, 124],

    # "MAN/man-32-40-imo-tier-iii-marine.pdf",
    [2, 87],
    [2, 88],
    [2, 89],
    [2, 92],
    [2, 93],
    [2, 98],
    [2, 99],
    [2, 100],
    [2, 101],
    [2, 102],
    [2, 103],
    [2, 104],
    [2, 105],
    [2, 106],
    [2, 107],
    [2, 108],
    [2, 109],
    [2, 110],
    [2, 111],
    [2, 112],
    [2, 113],
    [2, 114],
    [2, 115],
    [2, 116],
    [2, 117],
    [2, 118],
    [2, 119],
    [2, 120],
    [2, 121],
    [2, 122],
    [2, 123],
    [2, 124],
    [2, 125],
    [2, 126],
    [2, 127],
    [2, 128],
    [2, 129],
    [2, 130],
    [2, 131],
    [2, 132],
    [2, 133],
    [2, 134],
    
    # "MAN/man-32-44cr-imo-tier-ii-marine.pdf",
    [3, 77],
    [3, 78],
    [3, 79],
    [3, 80],
    [3, 81],
    [3, 82],
    [3, 84],
    [3, 89],
    [3, 90],
    [3, 91],
    [3, 92],
    [3, 93],
    [3, 94],
    [3, 95],
    [3, 96],
    [3, 97],
    [3, 98],
    [3, 99],
    [3, 100],
    [3, 101],
    [3, 102],
    [3, 103],
    [3, 104],
    [3, 105],
    [3, 106],
    [3, 107],
    [3, 108],
    [3, 109],
    [3, 110],
    [3, 111],
    [3, 112],
    [3, 113],
    [3, 114],
    [3, 115],
    [3, 116],
    [3, 117],
    [3, 118],
    [3, 119],
    [3, 120],
    [3, 121],
    [3, 122],
    [3, 123],
    [3, 124],
    [3, 125],
    [3, 126],
    [3, 127],
    [3, 128],

    # "MAN/man-32-44cr-imo-tier-iii-marine.pdf",
    [4, 88],
    [4, 89],
    [4, 90],
    [4, 91],
    [4, 92],
    [4, 93],
    [4, 95],
    [4, 101],
    [4, 102],
    [4, 103],
    [4, 104],
    [4, 105],
    [4, 106],
    [4, 107],
    [4, 108],
    [4, 109],
    [4, 110],
    [4, 111],
    [4, 112],
    [4, 113],
    [4, 114],
    [4, 115],
    [4, 116],
    [4, 117],
    [4, 118],
    [4, 119],
    [4, 120],
    [4, 121],
    [4, 122],
    [4, 123],
    [4, 124],
    [4, 125],
    [4, 126],
    [4, 127],
    [4, 128],
    [4, 129],
    [4, 130],
    [4, 131],
    [4, 132],
    [4, 133],
    [4, 134],
    [4, 135],
    [4, 136],
    [4, 137],
    [4, 138],
    [4, 139],
    [4, 140],
    
    # "MAN/man-48-60cr-imo-tier-ii-marine.pdf",
    [5, 70],
    [5, 71],
    [5, 72],
    [5, 73],
    [5, 75],
    [5, 79],
    [5, 80],
    [5, 81],
    [5, 82],
    [5, 83],
    [5, 84],
    [5, 85],
    [5, 86],
    [5, 87],
    [5, 88],
    [5, 89],
    [5, 90],
    [5, 91],
    [5, 92],
    [5, 93],
    [5, 94],
    [5, 95],
    [5, 96],
    [5, 97],
    [5, 98],
    [5, 99],
    [5, 100],
    [5, 101],
    [5, 102],
    [5, 103],
    [5, 104],
    [5, 105],
    [5, 106],
    [5, 107],
    [5, 108],
    [5, 109],
    [5, 110],
    [5, 111],
    [5, 112],
    [5, 113],
    [5, 114],
    [5, 115],
    [5, 116],
    
    # "MAN/man-48-60cr-imo-tier-iii-marine.pdf",
    [6, 79],
    [6, 80],
    [6, 81],
    [6, 82],
    [6, 84],
    [6, 89],
    [6, 90],
    [6, 91],
    [6, 92],
    [6, 93],
    [6, 94],
    [6, 95],
    [6, 96],
    [6, 97],
    [6, 98],
    [6, 99],
    [6, 100],
    [6, 101],
    [6, 102],
    [6, 103],
    [6, 104],
    [6, 105],
    [6, 106],
    [6, 107],
    [6, 108],
    [6, 109],
    [6, 110],
    [6, 111],
    [6, 112],
    [6, 113],
    [6, 114],
    [6, 115],
    [6, 116],
    [6, 117],
    [6, 118],
    [6, 119],
    [6, 120],
    [6, 121],
    [6, 122],
    [6, 123],
    [6, 124],
    [6, 125],
    [6, 126],
    
    # "MAN/man-51-60df-imo-tier-ii-imo-tier-iii-marine.pdf",
    [7, 78],
    [7, 79],
    [7, 80],
    [7, 81],
    [7, 82],
    [7, 83],
    [7, 86],
    [7, 90],
    [7, 91],
    [7, 92],
    [7, 93],
    [7, 94],
    [7, 95],
    [7, 96],
    [7, 97],
    [7, 98],
    [7, 99],
    [7, 100],
    [7, 101],
    [7, 102],
    [7, 103],
    [7, 104],
    [7, 105],
    [7, 106],
    [7, 107],
    [7, 108],
    [7, 109],
    [7, 110],
    [7, 111],
    [7, 112],
    [7, 113],
    [7, 114],
    [7, 115],
    [7, 116],
    [7, 118],
    [7, 119],
    [7, 120],
    [7, 121],
    [7, 122],
    [7, 123],
    [7, 124],
    [7, 125],
    [7, 126],
    [7, 127],
    [7, 128],
    [7, 130],
    [7, 131],
    [7, 132],
    [7, 133],
    [7, 134],
    [7, 135],
    [7, 136],
    [7, 137],
    
    # "MAN/man-51-60df-imo-tier-iii-marine.pdf",
    [8, 87],
    [8, 88],
    [8, 89],
    [8, 90],
    [8, 91],
    [8, 92],
    [8, 94],
    [8, 101],
    [8, 102],
    [8, 103],
    [8, 104],
    [8, 105],
    [8, 106],
    [8, 107],
    [8, 108],
    [8, 109],
    [8, 110],
    [8, 111],
    [8, 112],
    [8, 113],
    [8, 114],
    [8, 115],
    [8, 116],
    [8, 117],
    [8, 118],
    [8, 119],
    [8, 120],
    [8, 121],
    [8, 122],
    [8, 123],
    [8, 124],
    [8, 125],
    [8, 126],
    [8, 127],
    [8, 129],
    [8, 130],
    [8, 131],
    [8, 132],
    [8, 133],
    [8, 134],
    [8, 135],
    [8, 136],
    [8, 137],
    [8, 138],
    [8, 139],
    [8, 141],
    [8, 142],
    [8, 143],
    [8, 144],
    [8, 145],
    [8, 146],
    [8, 147],
    [8, 148],
    
    # "MAN/man-es_l3_factsheets_4p_man_175d_genset_n5_201119_low_new.pdf",
    [9, 2],
    [9, 3],
    
    # "MAN/man-l32-44-genset-imo-tier-ii-marine.pdf",
    [10, 57],
    [10, 62],
    [10, 63],
    [10, 64],
    [10, 65],
    [10, 66],
    
    # "MAN/man-l32-44-genset-imo-tier-iii-marine.pdf",
    [11, 67],
    [11, 73],
    [11, 74],
    [11, 75],
    [11, 76],
    [11, 77],
    
    # "MAN/man-l35-44df-imo-tier-ii-marine.pdf",
    [12, 78],
    [12, 79],
    [12, 80],
    [12, 81],
    [12, 82],
    [12, 84],
    [12, 89],
    [12, 90],
    [12, 91],
    [12, 92],
    [12, 93],
    [12, 94],
    [12, 95],
    [12, 96],
    [12, 97],
    [12, 98],
    [12, 99],
    [12, 100],
    [12, 101],
    [12, 102],
    [12, 103],
    [12, 104],
    [12, 105],
    [12, 106],
    [12, 107],
    [12, 108],
    [12, 109],
    [12, 110],
    [12, 111],
    [12, 112],
    [12, 113],
    
    # "MAN/man-l35-44df-imo-tier-iii-marine.pdf",
    [13, 88],
    [13, 89],
    [13, 90],
    [13, 91],
    [13, 92],
    [13, 94],
    [13, 101],
    [13, 102],
    [13, 103],
    [13, 104],
    [13, 105],
    [13, 106],
    [13, 107],
    [13, 108],
    [13, 109],
    [13, 110],
    [13, 111],
    [13, 112],
    [13, 113],
    [13, 114],
    [13, 115],
    [13, 116],
    [13, 117],
    [13, 118],
    [13, 119],
    [13, 120],
    [13, 121],
    [13, 122],
    [13, 123],
    [13, 124],
    [13, 125],
    
    # "MAN/man-v28-33d-stc-imo-tier-ii-marine.pdf",
    [14, 56],
    [14, 57],
    [14, 58],
    [14, 59],
    [14, 63],
    [14, 64],
    [14, 65],
    [14, 66],
    [14, 67],
    [14, 68],
    [14, 69],
    [14, 70],
    [14, 71],
    
    # "MAN/man-v28-33d-stc-imo-tier-iii-marine.pdf",
    [15, 66],
    [15, 67],
    [15, 68],
    [15, 69],
    [15, 74],
    [15, 75],
    [15, 76],
    [15, 77],
    [15, 78],
    [15, 79],
    [15, 80],
    [15, 81],
    [15, 82],
    
    # "MAN/PG_M-III_L2131.pdf",
    [16, 35],
    [16, 37],
    [16, 139],
    
    # "MAN/PG_M-II_L1624.pdf",
    [17, 35],
    [17, 37],
    [17, 141],
    
    # "MAN/PG_P-III_L2131.pdf", NO DATA
    
    # "MAN/PG_P-II_L2131.pdf",
    [19, 105],
    [19, 106],
    [19, 197],
    
    # "MAN/PG_P-I_L2131.pdf", NO DATA
    
    # "Wartsila/brochure-o-e-w31.pdf", NO DATA
    
    # "Wartsila/brochure-o-e-wartsila-dual-fual-low-speed.pdf", NO DATA
    
    # "Wartsila/leaflet-w14-2018.pdf", NO DATA
    
    # "Wartsila/marine-power-catalogue.pdf", NO DATA
    
    # "Wartsila/product-guide-o-e-w26.pdf",
    [25, 22],
    [25, 23],
    [25, 25],
    [25, 26],
    [25, 28],
    [25, 29],
    [25, 31],
    [25, 32],
    [25, 34],
    [25, 35],
    [25, 36],
    [25, 37],
    [25, 38],
    [25, 39],
    [25, 40],
    [25, 42],
    [25, 43],
    [25, 45],
    [25, 46],
    [25, 48],
    [25, 49],
    [25, 50],
    [25, 51],
    [25, 52],
    [25, 53],
    [25, 54],
    [25, 56],
    [25, 57],
    [25, 59],
    [25, 60],
    [25, 62],
    [25, 63],
    
    # "Wartsila/product-guide-o-e-w31df.pdf", NO DATA
    
    # "Wartsila/product-guide-o-e-w32.pdf", NO DATA
    
    # "Wartsila/product-guide-o-e-w34df.pdf", NO DATA
    
    # "Wartsila/product-guide-o-e-w46df.pdf",
    [29, 25],
    [29, 26],
    [29, 27],
    [29, 29],
    [29, 30],
    [29, 31],
    [29, 33],
    [29, 34],
    [29, 35],
    [29, 37],
    [29, 38],
    [29, 39],
    [29, 41],
    [29, 42],
    [29, 43],
    [29, 45],
    [29, 46],
    [29, 47],
    [29, 49],
    [29, 50],
    [29, 51],

    # "Wartsila/product-guide-o-e-w46f.pdf",
    [30, 20],
    [30, 21],
    [30, 23],
    [30, 24],
    [30, 26],
    [30, 27],
    [30, 29],
    [30, 30],
    [30, 32],
    [30, 33],
    [30, 35],
    [30, 36],
    [30, 38],
    [30, 39],
    
    # "Wartsila/wartsila-14-product-guide.pdf", NO DATA
    
    # Wartsila/wartsila-31sg-product-guide.pdf", NO DATA
    
    # "Wartsila/wartsila-46df.pdf", NO DATA
    
    # "Wartsila/wartsila-o-e-ls-x92.pdf",
    [34, 2],

    # "Hyundai_HHI/2016_HHI_en.pdf", NO DATA
    
    # "Hyundai_HHI/2017_HHI_en.pdf", NO DATA
    
    # "Hyundai_HHI/2018_HHI_en.pdf", NO DATA
    
    # "Hyundai_HHI/2019_HHI_en.pdf", NO DATA
    
    # "Hyundai_HHI/2020_KSOE_IR_EN.pdf", NO DATA
    
    # "Hyundai_HHI/2021_KSOE_IR_EN.pdf", NO DATA
    
    # "Hyundai_HHI/HHI_EMD_brochure2017_1.pdf",
    [41, 10],

    # "Hyundai_HHI/HHI_EMD_brochure2017_2.pdf", NO DATA
    
    # "Hyundai_HHI/HHI_EMD_brochure2017_3.pdf", NO DATA
    
    # "Hyundai_HHI/HHI_EMD_brochure2019_1.pdf",
    [44, 9],

    # "Hyundai_HHI/speclal_NSD2020.pdf", NO DATA

    # "WinGD/16-06-pres-kit_x92.pdf",
    [46, 6],
    
    # "WinGD/cimac2016_120_kyrtatos_paper_thedevelopmentofthemodern2strokemarinedieselengine.pdf", NO DATA
    
    # "WinGD/cimac2016_173_brueckl_paper_virtualdesign-and-simulation-in2strokemarineenginedevelopment.pdf", NO DATA
    
    # "WinGD/cimac2016_233_ott_paper_x-df.pdf", NO DATA
    
    # "WinGD/dominikschneiter_tieriii-programme.pdf", NO DATA
    
    # "WinGD/Fuel-Flexible-Injection-System-How-to-Handle-a-Fuel-Specturm-from-Diesel_CIMAC2019_paper_404_Andreas-Schmid.pdf", NO DATA
    
    # "WinGD/marcspahni_generation-x-engines.pdf", NO DATA

    #"WinGD/MIM_WinGD-RT-flex48T-D.pdf",
    [53, 10],
    [53, 64],

    # "WinGD/MIM_WinGD-X72.pdf",
    [54, 19],
    
    # "WinGD/MM_WinGD-RT-flex58T-D.pdf", NO DATA
    
    # "WinGD/MM_WinGD-X35-B_2021-09.pdf", NO DATA
    
    # "WinGD/MM_WinGD-X52.pdf",
    [57, 1046],

    # "WinGD/MM_WinGD-X72.pdf", NO DATA
    
    # "WinGD/MM_WinGD-X72DF.pdf", NO DATA
    
    # "WinGD/MM_WinGD-X82-B.pdf", NO DATA
    
    # "WinGD/Motorship-May-2018-VP-R-D-Leader-Brief.pdf", NO DATA
    
    # "WinGD/OM_WinGD-X62_2021-09.pdf",
    [62, 509],

    # "WinGD/OM_WinGD-X72DF_2021-09.pdf",
    [63, 583],

    # "WinGD/OM_WinGD-X82-B.pdf",
    [64, 449],

    # "WinGD/OM_WinGD_RT-flex50DF.pdf",
    [65, 567],

    # "WinGD/WinGD-12X92DF-Development-of-the-Most-Powerful-Otto-Engine-Ever_CIMAC2019_paper_425_Dominik-Schneiter.pdf", NO DATA
    
    # "WinGD/wingd-paper_engine_selection_for_very_large_container_vessels_201609.pdf",
    [67, 9],

    # "WinGD/WinGD-WiDE-Brochure.pdf", NO DATA
    
    # "WinGD/WinGD_Engine-Booklet_2021.pdf",
    [69, 12],
    [69, 13],
    [69, 14],
    [69, 15],
    [69, 16],
    [69, 17],
    [69, 18],
    [69, 19],
    [69, 20],
    [69, 22],
    [69, 23],
    [69, 24],
    [69, 25],
    [69, 26],
    [69, 27],

    # "WinGD/wingd_moving-inlet-ports-concept-for-optimization-of-2-stroke-uni-flow-engines_patrick-rebecchi.pdf", NO DATA

    # "WinGD/X-DF-Engines-by-WinGD.pdf",
    [71, 4],

    # "WinGD/X-DF-FAQ.pdf", NO DATA
    
]

# Minimum and maximum document index by manufacturer. Note that manufacturer 1 (MAN) is in index 2 etc.
manufMinDoc = [-1, -1,  0, 21, 35, 46]
manufMaxDoc = [-1, -1, 20, 34, 45, 72]

# Pages found by the program: Doc number, page number. Phase 1 algorithm should find 707 pages
found_doc_pages = []

# All pages are divided into 3 sets: train set, validation/development set, test set
all_doc_pages_tvt = [] # { doc: <docNumber>, page: <pageNumber>, is_train: <in training set>, is_valid: <in validation set>, is_test: <in test set>}

# Vocabulary over all pages in all PDF files (=18 019 pages from 73 files)
vocab = Counter()

# Page counts of each doc
doc_page_counts = [0] * len(files_to_process)

# The training and validation sets
Xtrain = None
Ytrain = None
Xvalid = None
Yvalid = None

# The tokenizer, fitted to training data
tokenizer = None

# The Keras model, trained with training data
model = None

# Program mode: 0 = Train/Valid/Test sets; 1 = Use all samples for training and evaluation
#    2 = manufacturer 1: valid & test, 3 = manuf 2: v & t, 4 = manuf. 3: v & t, 5 = manuf. 4: v & t
modelUseSamples = 0

# NN Evaluation options
isConvertToLowercase = True
isRemovePunctuation = True
handleNumbers_RmKpRp = 0 # 0=Remove, 1=Keep, 2=Replace with 'WeHaveANumber'
isRemoveStopWords = True
minimumWordLength = 2
minVocabularyOccur = 2
scoreWordsMethod = "binary" # binary, count, tfidf, freq
duplicatePosSamples = 0
isUseTwoNNLayers = False
neuronsInNNLayer1 = 50
neuronsInNNLayer2 = 20


# The following 'global' functions have been copied from web site https://machinelearningmastery.com/deep-learning-bag-of-words-model-sentiment-analysis/.

# ****** Phase 1 ******

# turn a doc (=PDF page) into clean tokens
def clean_doc(doc):
    # split into tokens by white space
    tokens = doc.split()
    # turn all words into lowercase
    if isConvertToLowercase:
        tokens = [s.lower() for s in tokens]
    # remove punctuation from each token
    if isRemovePunctuation:
        table = str.maketrans('', '', string.punctuation)
        tokens = [w.translate(table) for w in tokens]
    # Handle numbers
    if handleNumbers_RmKpRp == 0:
        # remove remaining tokens that are not alphabetic
        tokens = [word for word in tokens if word.isalpha()]
    elif handleNumbers_RmKpRp == 1:
        # keep numbers = remove remaining tokens that are not alphanumeric
        tokens = [word for word in tokens if word.isalnum()]
    elif handleNumbers_RmKpRp == 2:
        tokens = [word for word in tokens if word.isalnum()]
        tokens = ['WeHaveANumber' if word.isnumeric() else word for word in tokens]
    # filter out stop words
    if isRemoveStopWords:
        stop_words = set(stopwords.words('english'))
        tokens = [w for w in tokens if not w in stop_words]
    # filter out short tokens
    if minimumWordLength>1:
        tokens = [word for word in tokens if len(word) >= minimumWordLength]
    return tokens


# ****** Phase 2 ******

# clean doc and add to vocab
def add_doc_to_vocab(page, vocab):
	# clean doc
	tokens = clean_doc(page)
	# update counts
	vocab.update(tokens)
 
# add all words in training set docs to vocabulary; SIDE EFFECTS: divide pages into training, validation, and test sets; store page counts
def process_pages_for_vocabulary(vocab, self_frame):
    global all_doc_pages_tvt
    global doc_page_counts
    
    for docNumber, filename in enumerate(files_to_process):
        # creating a pdf file object 
        doc = fitz.open(filename)
        self_frame.m_extracted.AppendText(filename+'\n')
        self_frame.m_extracted.Refresh()
        self_frame.m_extracted.Update()
        doc_page_counts[docNumber] = doc.page_count
        
        # creating a page object; ALSO determine if a page belongs to train/validation/test set
        for pageNumber, page in enumerate(doc):
            rnd10 = random.randrange(10)
            rnd_tvt = 1 if rnd10<=5 else (2 if rnd10<=7 else 3)
            is_train = rnd_tvt==1
            is_valid = rnd_tvt==2
            is_test  = rnd_tvt==3
            if modelUseSamples==1:
                is_train = True
                is_valid = True
                is_test  = True
            elif modelUseSamples>1:
                testDoc =  docNumber>=manufMinDoc[modelUseSamples] and docNumber<=manufMaxDoc[modelUseSamples]
                is_train = not testDoc
                is_valid = testDoc
                is_test  = testDoc
            all_doc_pages_tvt.append({ "doc": docNumber, "page": pageNumber+1, "is_train": is_train, "is_valid": is_valid, "is_test": is_test })
            if is_train:
                pageText = page.get_text('text')
                add_doc_to_vocab(pageText, vocab)

 
# (WAS INLINE CODE) define vocabulary
def create_vocabulary(self_frame):
    global vocab
    
    vocab = Counter()
    # add all docs to vocab
    process_pages_for_vocabulary(vocab, self_frame)
   # print the size of the vocab
    print('Vocabulary items:', len(vocab))
    # print the top words in the vocab
    print(vocab.most_common(50))
    if minVocabularyOccur>1:
        # keep tokens with a min occurrence
        tokens = [k for k,c in vocab.items() if c >= minVocabularyOccur]
        print('Vocabulary items, min_occ=', minVocabularyOccur, ':', len(tokens))
    else:
        tokens = [k for k,c in vocab.items()]
        print('Vocabulary items:', len(tokens))

    # Saving vocabulary to file 'vocab.txt' removed...
    vocab = tokens
    

# ****** Phase 3 ****** SKIPPED ******

# clean doc (=PDF page) and return line of tokens (tokens not in vocabulary removed)
def doc_to_line(doc, vocab):
	# clean doc
	tokens = clean_doc(doc)
	# filter by vocab
	tokens = [w for w in tokens if w in vocab]
	return ' '.join(tokens)

# load all docs in training set
def process_docs_get_train_page_lines(vocab, self_frame):
    lines = list()
    for docNumber, filename in enumerate(files_to_process):
        # creating a pdf file object 
        doc = fitz.open(filename)
        self_frame.m_extracted.AppendText(filename+'\n')
        self_frame.m_extracted.Refresh()
        self_frame.m_extracted.Update()
        
        # creating a page object; skip those pages that are not in the training set
        for pageNumber, page in enumerate(doc):
            traValTes = [x for x in all_doc_pages_tvt if x["doc"]==docNumber and x["page"]==pageNumber+1]
            if len(traValTes)>0 and traValTes[0]["is_train"]:
                pageText = page.get_text('text')
                # load and clean the doc
                line = doc_to_line(pageText, vocab)
                # add to list
                lines.append(line)
                
    return lines


# (WAS INLINE CODE) use vocabulary, load training set data
def load_training_set_data(self_frame):
    global vocab
    
    # load the vocabulary
    vocab = set(vocab)
    # load all training pages
    all_lines = process_docs_get_train_page_lines(vocab, self_frame)
    # summarize what we have
    print('Training set pages:', len(all_lines))


# ****** Phase 4 ******

# load all docs in training OR validation set
def process_docs_get_train_or_valid_page_lines(vocab, is_train, self_frame):
    lines = list()
    for docNumber, filename in enumerate(files_to_process):
        # creating a pdf file object 
        doc = fitz.open(filename)
        self_frame.m_extracted.AppendText(filename+'\n')
        self_frame.m_extracted.Refresh()
        self_frame.m_extracted.Update()
        
        # creating a page object; skip those pages that are not in the training (is_train==True) or validation (is_train==False) set
        for pageNumber, page in enumerate(doc):
            traValTes = [x for x in all_doc_pages_tvt if x["doc"]==docNumber and x["page"]==pageNumber+1]
            if len(traValTes)>0:
                if (is_train and traValTes[0]["is_train"]) or (not is_train and traValTes[0]["is_valid"]):
                    pageText = page.get_text('text')
                    # load and clean the doc
                    line = doc_to_line(pageText, vocab)
                    # add to list
                    lines.append(line)
                    # if we are duplicating positive samples, then do duplicate for training set ground truth samples
                    if duplicatePosSamples>0 and [docNumber, pageNumber+1] in ground_truth_doc_pages:
                        if is_train and traValTes[0]["is_train"]:
                            for d in range(duplicatePosSamples):
                                lines.append(line)
                
    return lines


# (WAS INLINE CODE) use vocabulary, load training and validation set data
def create_x_train_x_test(self_frame):
    global vocab
    global Xtrain
    global Xvalid
    global tokenizer
    
    # load all training reviews
    train_lines = process_docs_get_train_or_valid_page_lines(vocab, True, self_frame)
    
    # create the tokenizer
    tokenizer = Tokenizer()
    # fit the tokenizer on the documents
    tokenizer.fit_on_texts(train_lines)
    
    # encode training data set
    Xtrain = tokenizer.texts_to_matrix(train_lines, mode=scoreWordsMethod)
    print('Xtrain shape:', Xtrain.shape)
    
    # load all test reviews
    valid_lines = process_docs_get_train_or_valid_page_lines(vocab, False, self_frame)
    # encode training data set
    Xvalid = tokenizer.texts_to_matrix(valid_lines, mode=scoreWordsMethod)
    print('Xvalid shape:', Xvalid.shape)


# ****** Phase 5 ******

# (WAS INLINE CODE) use 'Xtrain' and 'Xvalid'; compute 'Ytrain' and 'Yvalid'; create and evaluate Keras model
def create_xy_train_valid(self_frame):
    global Xtrain
    global Ytrain
    global Xvalid
    global Yvalid
    global model

    Ytrain = array([0] * Xtrain.shape[0]) # array([0 for _ in range(Xtrain.shape[0])]) # 
    Yvalid = array([0] * Xvalid.shape[0]) # array([0 for _ in range(Xvalid.shape[0])]) # 

    idx_train = 0
    idx_valid = 0    
    for docNumber, filename in enumerate(files_to_process):
        # creating a pdf file object 
        self_frame.m_extracted.AppendText(filename+'\n')
        self_frame.m_extracted.Refresh()
        self_frame.m_extracted.Update()
        
        # creating a page object; ALSO determine if a page belongs to train/validation/test set
        for pageNumber in range(doc_page_counts[docNumber]):
            traValTes = [x for x in all_doc_pages_tvt if x["doc"]==docNumber and x["page"]==pageNumber+1]
            if len(traValTes)>0:
                if traValTes[0]["is_train"]: 
                    if [docNumber, pageNumber+1] in ground_truth_doc_pages:
                        Ytrain[idx_train] = 1
                        if duplicatePosSamples>0:
                            for d in range(duplicatePosSamples):
                                idx_train += 1
                                Ytrain[idx_train] = 1
                    idx_train += 1
                if traValTes[0]["is_valid"]: 
                    if [docNumber, pageNumber+1] in ground_truth_doc_pages:
                        Yvalid[idx_valid] = 1
                    idx_valid += 1
    
    n_words = Xvalid.shape[1]
    # define network
    model = Sequential()
    model.add(Dense(neuronsInNNLayer1, input_shape=(n_words,), activation='relu'))
    if isUseTwoNNLayers:
        model.add(Dense(neuronsInNNLayer2, activation='relu'))
    model.add(Dense(1, activation='sigmoid'))
    # compile network
    model.compile(loss='binary_crossentropy', optimizer='adam', metrics=['accuracy'])
    # fit network
    model.fit(Xtrain, Ytrain, epochs=50, verbose=2)
    # evaluate
    loss, acc = model.evaluate(Xvalid, Yvalid, verbose=0)
    print('Validation Accuracy: %f %%' % (acc*100))


# ****** Phase 6 ******

# evaluate a neural network model
def evaluate_mode(Xtrain, Ytrain, Xvalid, Yvalid):
    scores = list()
    n_repeats = 6 # WAS: 30
    n_words = Xvalid.shape[1]
    for i in range(n_repeats):
        # define network
        model = Sequential()
        model.add(Dense(neuronsInNNLayer1, input_shape=(n_words,), activation='relu'))
        if isUseTwoNNLayers:
            model.add(Dense(neuronsInNNLayer2, activation='relu'))
        model.add(Dense(1, activation='sigmoid'))
        # compile network
        model.compile(loss='binary_crossentropy', optimizer='adam', metrics=['accuracy'])
        # fit network
        model.fit(Xtrain, Ytrain, epochs=50, verbose=2)
        # evaluate
        loss, acc = model.evaluate(Xvalid, Yvalid, verbose=0)
        scores.append(acc)
        print('Repeat %d/%d => Accuracy: %s\n' % ((i+1), n_repeats, acc))
        return scores

# prepare bag of words encoding of docs
def prepare_data(train_docs, valid_docs, mode):
    global tokenizer

    # create the tokenizer
    tokenizer = Tokenizer()
    # fit the tokenizer on the documents
    tokenizer.fit_on_texts(train_docs)
    # encode training data set
    Xtrain = tokenizer.texts_to_matrix(train_docs, mode=mode)
    # encode training data set
    Xvalid = tokenizer.texts_to_matrix(valid_docs, mode=mode)
    return Xtrain, Xvalid


# (WAS INLINE CODE) use vocabulary, load training set data   ===>   TODO: need global variables 'Xtrain' and 'Xvalid'
def evaluate_modes(self_frame):
    global vocab
    global Ytrain
    global Yvalid

    # load all training reviews
    train_lines = process_docs_get_train_or_valid_page_lines(vocab, True, self_frame)
    # load all test reviews
    valid_lines = process_docs_get_train_or_valid_page_lines(vocab, False, self_frame)
    
    modes = ['binary', 'count', 'tfidf', 'freq']
    results = DataFrame()
    for mode in modes:
    	print('\nEVALUATING MODE: %s' % (mode))
    	# prepare data for mode
    	Xtrain, Xvalid = prepare_data(train_lines, valid_lines, mode)
    	# evaluate model on data for mode
    	results[mode] = evaluate_mode(Xtrain, Ytrain, Xvalid, Yvalid)
    # summarize results
    print(results.describe())
    # plot results
    results.boxplot()
    pyplot.show()



# ****** Phase 7 -- USE MODEL ******
# ****** NOTE: Either Phase 6 OR Phase 7, not both ******

# evaluate a neural network model
# classify a page as negative (0) or positive (1)
def predict_page(page, vocab, tokenizer, model):
	# clean
	tokens = clean_doc(page)
	# filter by vocab
	tokens = [w for w in tokens if w in vocab]
	# convert to line
	line = ' '.join(tokens)
	# encode
	encoded = tokenizer.texts_to_matrix([line], mode=scoreWordsMethod)
	# prediction
	yhat = model.predict(encoded, verbose=0)
	return round(yhat[0,0])


# evaluate NN model and compare result to Ground Truth
def evaluate_test_set(self_frame):
    global vocab
    global tokenizer
    global model

    output_file = resultsFolder1 + "/NN_eval_results.txt"
    with open(output_file, 'w') as fp:
            
        nbr_true_pos = 0
        nbr_fals_pos = 0
        nbr_true_neg = 0
        nbr_fals_neg = 0
        nbr_test_smp = 0
        for docNumber, filename in enumerate(files_to_process):
            # creating a pdf file object 
            doc = fitz.open(filename)
            self_frame.m_extracted.AppendText(filename+'\n')
            self_frame.m_extracted.Refresh()
            self_frame.m_extracted.Update()
            
            # creating a page object for each page in the test set; evaluate NN model and compare result to Ground Truth
            for pageNumber, page in enumerate(doc):
                traValTes = [x for x in all_doc_pages_tvt if x["doc"]==docNumber and x["page"]==pageNumber+1]
                if len(traValTes)<=0 or not traValTes[0]["is_test"]:
                    continue
                nbr_test_smp += 1
                pageText = page.get_text('text')
                is_pos_ev = (predict_page(pageText, vocab, tokenizer, model) > 0)
                is_pos_gt = ([docNumber, pageNumber+1] in ground_truth_doc_pages)
                if is_pos_ev and is_pos_gt:
                    nbr_true_pos += 1
                elif not is_pos_ev and not is_pos_gt:
                    nbr_true_neg += 1
                elif is_pos_ev and not is_pos_gt:
                    nbr_fals_pos += 1
                elif not is_pos_ev and is_pos_gt:
                    nbr_fals_neg += 1
                    
                if is_pos_ev:
                    pix = page.get_pixmap()
                    png_filename = resultsFolder1 + "/Doc" + str(docNumber) + "_page" + str(pageNumber+1) + ".png"
                    pix.save(png_filename)
                    url = "file:///" + os.path.abspath(filename) + "#page=" + str(pageNumber+1)
                    print(url, file=fp)
                    # Occasionally headers contain non-printing characters; print what can be printed, skip the rest
                    try:
                        print(findTocItems(doc, pageNumber+1), file=fp)
                    except:
                        toc_list = findTocItems(doc, pageNumber+1)
                        print("['", end='', file=fp)
                        for toc in toc_list:
                            try:
                                print(toc, end="', '", file=fp)
                            except:
                                for t in toc:
                                    try:
                                        print(t, end='', file=fp)
                                    except:
                                        pass
                                print("', '", end='', file=fp)
                        print("']", file=fp)
                    
                    # extracting html from page
                    pageHtml = page.get_text('html', sort=True)
                    html_filename = resultsFolder1 + "/Doc" + str(docNumber) + "_page" + str(pageNumber+1) + ".html"
                    html_file = open(html_filename, "w")
                    html_file.write(pageHtml)
                    html_file.close()
                    
                    # 200 characters per line AND separate text line for each different "top" coord in HTML,
                    # If we make value 200 (chars) much smaller, columns are no longer separated by enough space chars.
                    (wid, hei) = findPageInfo(pageHtml)
                    xwid = int(round(wid/200))
                    tops = countLinesInPage(pageHtml)
                    nwid = int(round(wid/xwid))
                    nhei = len(tops)
                    linesToPrint = findMatchingLines(pageHtml, xwid, tops, nwid, nhei)
                        
                    self_frame.m_extracted.AppendText("   Doc:"+str(docNumber)+", Page:"+str(pageNumber+1)+
                                                      (" (BOGUS = False positive)\n" if not is_pos_gt else "\n"))
                    print("Doc:"+str(docNumber)+", Page:"+str(pageNumber+1)+" => ", file=fp)
                    
                    for ltp in linesToPrint:
                        try:
                            print(ltp.rstrip(), file=fp)
                        except:
                            pass
                elif not is_pos_ev and is_pos_gt:
                    self_frame.m_extracted.AppendText("   Doc:"+str(docNumber)+", Page:"+str(pageNumber+1)+" (MISSING = False negative)\n")
        
    self_frame.m_extracted.AppendText(  'nbr_true_pos: ' + str(nbr_true_pos) +
                                      ', nbr_fals_pos: ' + str(nbr_fals_pos) +
                                      ', nbr_true_neg: ' + str(nbr_true_neg) +
                                      ', nbr_fals_neg: ' + str(nbr_fals_neg) + '\n')
    self_frame.m_extracted.Refresh()
    self_frame.m_extracted.Update()
    if nbr_test_smp != (nbr_true_pos+nbr_fals_pos+nbr_true_neg+nbr_fals_neg):
        print('ERROR: nbr_test_smp!=sum():', nbr_test_smp, (nbr_true_pos+nbr_fals_pos+nbr_true_neg+nbr_fals_neg))
    print('nbr_true_pos, nbr_fals_pos, nbr_true_neg, nbr_fals_neg:', nbr_true_pos, nbr_fals_pos, nbr_true_neg, nbr_fals_neg)
    return (nbr_true_pos, nbr_fals_pos, nbr_true_neg, nbr_fals_neg)

    
# ****** End of BagOfWords example ******



def findTocItems(doc, page):
    toc = doc.get_toc()
    if not toc:
        return []

    prev_page = 0
    items = []
    for level in range(1,6):
        curr_page = 0
        curr_text = ""
        for t in toc:
            if t[0]==level and t[2]>=prev_page and t[2]<=page:
                curr_page = t[2]
                curr_text = html.unescape(unicodedata.normalize('NFKC', t[1]))
            elif t[0]==level and t[2]>page:
                if curr_text:
                    items.append(curr_text)
                    prev_page = curr_page
                break
    return items


def findLineInfo(ln):
    ispan = ln.find("<span")
    if ispan<0:
        return [-1, -1, ""]
    iend = ln.find(">", ispan)
    iesp = ln.find("</span", iend)
    text = ln[iend+1:iesp]
    itop = ln.find("top:")
    ietp = ln.find("pt", itop)
    ileft = ln.find("left:")
    ielf = ln.find("pt", ileft)
    top = int(float(ln[itop+4:ietp]))
    left = int(float(ln[ileft+5:ielf]))
    return [top,left,text]


def findPageInfo(pg):
    idiv = pg.find("<div")
    if idiv<0:
        return [-1, -1]
    iend = pg.find(">", idiv)
    iwid = pg.find("width:", idiv)
    iewd = pg.find("pt", iwid)
    wid = int(float(pg[iwid+6:iewd]))
    ihei = pg.find("height:", idiv)
    iehe = pg.find("pt", ihei)
    hei = int(float(pg[ihei+7:iehe]))
    return [wid,hei]
    
    
def findMatchingLines(pageHtml, xwid, line_tops, nwid, nhei):
    spaces200 = " " * nwid
    linesToPrint = [spaces200] * nhei
    lines = pageHtml.split('\n')
    minleft = nwid*xwid
    for ln in lines:
        [top, left, text] = findLineInfo(ln)
        if top<0 or left<0:
            continue
        if left<minleft:
            minleft = left
    minleft /= xwid
    minleft -= 1
    
    for ln in lines:
        [top, left, text] = findLineInfo(ln)
        if top<0 or left<0:
            continue
        y = line_tops.index(top)
        x = int(round(left/xwid)-minleft)
        try:
            text = html.unescape(unicodedata.normalize('NFKC', text))
        except:
            pass
        linesToPrint[y] = linesToPrint[y][0:x] + text + linesToPrint[y][x+len(text):]
        
    return linesToPrint


def countLinesInPage(pageHtml):
    tops = set()
    lines = pageHtml.split('\n')
    for ln in lines:
        [top, left, text] = findLineInfo(ln)
        if top<0 or left<0:
            continue
        tops.add(top)
    s_tops = sorted(tops)
    return s_tops
        


class HtmlWindow(wx.html.HtmlWindow):
    def __init__(self, parent, id, size=(600,400)):
        wx.html.HtmlWindow.__init__(self,parent, id, size=size)
        if "gtk2" in wx.PlatformInfo:
            self.SetStandardFonts()

    def OnLinkClicked(self, link):
        wx.LaunchDefaultBrowser(link.GetHref())


        
class AboutBox(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, None, -1, "About VesselAI_ExtractData_Neural_Network",
            style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|
                wx.TAB_TRAVERSAL)
        hwin = HtmlWindow(self, -1, size=(400,200))
        vers = {}
        vers["python"] = sys.version.split()[0]
        vers["wxpy"] = wx.VERSION_STRING
        hwin.SetPage(aboutText % vers)
        btn = hwin.FindWindowById(wx.ID_OK)
        irep = hwin.GetInternalRepresentation()
        hwin.SetSize((irep.GetWidth()+25, irep.GetHeight()+10))
        self.SetClientSize(hwin.GetSize())
        self.CentreOnParent(wx.BOTH)
        self.SetFocus()



class Frame(wx.Frame):
    def __init__(self, title):
        wx.Frame.__init__(self, None, title=title, pos=(150,150), size=(1200,800))
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        menuBar = wx.MenuBar()
        menu = wx.Menu()
        m_exit = menu.Append(wx.ID_EXIT, "E&xit\tAlt-X", "Close window and exit program.")
        self.Bind(wx.EVT_MENU, self.OnClose, m_exit)
        menuBar.Append(menu, "&File")
        menu = wx.Menu()
        m_about = menu.Append(wx.ID_ABOUT, "&About", "Information about this program")
        self.Bind(wx.EVT_MENU, self.OnAbout, m_about)
        menuBar.Append(menu, "&Help")
        self.SetMenuBar(menuBar)
        
        self.statusbar = self.CreateStatusBar()

        self.m_docPage = ""

        splitter = wx.SplitterWindow(self, style=wx.SP_NO_XP_THEME | wx.SP_3D | wx.SP_LIVE_UPDATE)
        panel = wx.Panel(splitter)
        vert_box = wx.BoxSizer(wx.VERTICAL)
        hor_box_search = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_refresh = wx.BoxSizer(wx.HORIZONTAL)
        
        hor_box_programMode = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_processNumbers = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_scoreWords = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_shortWords = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_lowercase = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_removePuncts = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_removeStopWs = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_NN_parameters = wx.BoxSizer(wx.HORIZONTAL)
        
        m_textR0 = wx.StaticText(panel, -1, "Select the desired options for NN processing, then press button 'Evaluate', " + 
                                "and the drop-down below is filled with matching data pages.\n" +
                                "The standard procedure is to use separate Training/Validation/Test sets to evaluate the performance of the model,\n" +
                                "but after you have done that, you can train and evaluate the NN using ALL samples (for comparison).")
        vert_box.Add(m_textR0, 0, wx.ALL, 6)
        
        m_textM1 = wx.StaticText(panel, -1, "Program mode? (Train/Valid/Test sets vs.\nUse ALL data for training and evaluation)")
        hor_box_programMode.Add(m_textM1, 0, wx.ALL, 6)        
        self.m_evaluateSamples = wx.Choice(panel, choices = ['Evaluate Training/Validation/Test sets (recommended)', 
                                                         'Train and Evaluate ALL samples',
                                                         'Train with 3 manufacturers, test with manuf. #1 (MAN)',
                                                         'Train with 3 manufacturers, test with manuf. #2 (Wärtsilä)',
                                                         'Train with 3 manufacturers, test with manuf. #3 (HHI)',
                                                         'Train with 3 manufacturers, test with manuf. #4 (WinGD)'], size=(350,-1))
        self.m_evaluateSamples.Bind(wx.EVT_CHOICE, self.OnEvaluateSamples)
        self.m_evaluateSamples.SetSelection(0)
        hor_box_programMode.Add(self.m_evaluateSamples, 0, wx.ALL, 6)
        vert_box.Add(hor_box_programMode, 0, wx.ALL, 6)
        
        m_textM1 = wx.StaticText(panel, -1, "How to process numbers?\n(Remove/Keep/=>'WeHaveANumber')")
        hor_box_processNumbers.Add(m_textM1, 0, wx.ALL, 6)        
        self.m_processNumbers = wx.Choice(panel, choices = ['Remove numbers', 'Keep numbers', 'Replace numbers with fixed text'], size=(200,-1))
        self.m_processNumbers.Bind(wx.EVT_CHOICE, self.OnProcessNumbers)
        self.m_processNumbers.SetSelection(0)
        hor_box_processNumbers.Add(self.m_processNumbers, 0, wx.ALL, 6)
        
        m_textN1 = wx.StaticText(panel, -1, "How to score words?\n(4 methods of Keras Tokenizer)")
        hor_box_processNumbers.Add(m_textN1, 0, wx.ALL, 6)        
        self.m_bagOfWordsMethod = wx.Choice(panel, choices = ['binary', 'count', 'tfidf', 'freq'], size=(80,-1))
        self.m_bagOfWordsMethod.Bind(wx.EVT_CHOICE, self.OnBagOfWordsMethod)
        self.m_bagOfWordsMethod.SetSelection(0)
        hor_box_processNumbers.Add(self.m_bagOfWordsMethod, 0, wx.ALL, 6)
        vert_box.Add(hor_box_processNumbers, 0, wx.ALL, 6)

        m_textC1 = wx.StaticText(panel, -1, "Convert all words to lowercase?\n(i.e. remove case info)")
        hor_box_lowercase.Add(m_textC1, 0, wx.ALL, 6)        
        self.m_isLowercase = wx.CheckBox(panel, label="Turn to\nLowercase")        
        self.Bind(wx.EVT_CHECKBOX, self.OnIsLowercase) 
        self.m_isLowercase.SetValue(True)
        hor_box_lowercase.Add(self.m_isLowercase, 0, wx.ALL, 6)
        
        m_textP1 = wx.StaticText(panel, -1, "Remove Punctuation from all words?\n(commas ',', periods '.' etc.)")
        hor_box_lowercase.Add(m_textP1, 0, wx.ALL, 6)        
        self.m_isRemovePuncts = wx.CheckBox(panel, label="Remove\nPunctuations")        
        self.Bind(wx.EVT_CHECKBOX, self.OnIsRemovePunctuations) 
        self.m_isRemovePuncts.SetValue(True)
        hor_box_lowercase.Add(self.m_isRemovePuncts, 0, wx.ALL, 6)
        
        m_textS1 = wx.StaticText(panel, -1, "Remove English stop words from\nVocabulary? (e.g. 'a', 'an', 'the', 'of')")
        hor_box_lowercase.Add(m_textS1, 0, wx.ALL, 6)        
        self.m_isRemoveStopWs = wx.CheckBox(panel, label="Remove\nStop Words")        
        self.Bind(wx.EVT_CHECKBOX, self.OnIsRemoveStopWords) 
        self.m_isRemoveStopWs.SetValue(True)
        hor_box_lowercase.Add(self.m_isRemoveStopWs, 0, wx.ALL, 6)
        vert_box.Add(hor_box_lowercase, 0, wx.ALL, 6)
        
        m_textE1 = wx.StaticText(panel, -1, "Eliminate short words from Vocabulary?\n(shorter than minimum length; 1=No)")
        hor_box_shortWords.Add(m_textE1, 0, wx.ALL, 6)        
        self.m_textShortLimit = wx.StaticText(panel, -1, "Minimum\nword length")
        hor_box_shortWords.Add(self.m_textShortLimit, 0, wx.ALL, 6)
        self.m_shortLimit = wx.TextCtrl(panel, -1, "2", size=(50,-1))
        hor_box_shortWords.Add(self.m_shortLimit, 0, wx.ALL, 6)
        
        m_textE2 = wx.StaticText(panel, -1, "Eliminate infrequent words from Vocabulary?\n(count less than min occurrence; 1=No)")
        hor_box_shortWords.Add(m_textE2, 0, wx.ALL, 6)        
        self.m_textMinOccur = wx.StaticText(panel, -1, "Minimum\noccurrence")
        hor_box_shortWords.Add(self.m_textMinOccur, 0, wx.ALL, 6)
        self.m_minOccurrence = wx.TextCtrl(panel, -1, "2", size=(50,-1))
        hor_box_shortWords.Add(self.m_minOccurrence, 0, wx.ALL, 6)
        vert_box.Add(hor_box_shortWords, 0, wx.ALL, 6)
        
        m_textL3 = wx.StaticText(panel, -1, "Duplicate positive samples N times?\n(more weight to the rare positives; 0=No)")
        hor_box_NN_parameters.Add(m_textL3, 0, wx.ALL, 6)        
        self.m_textDuplPos = wx.StaticText(panel, -1, "Duplicate\nx N")
        hor_box_NN_parameters.Add(self.m_textDuplPos, 0, wx.ALL, 6)
        self.m_duplicatePos = wx.TextCtrl(panel, -1, "0", size=(50,-1))
        hor_box_NN_parameters.Add(self.m_duplicatePos, 0, wx.ALL, 6)
        m_textL0 = wx.StaticText(panel, -1, "Neural Network structure?\n(1 or 2 hidden layers, nbr neurons)")
        hor_box_NN_parameters.Add(m_textL0, 0, wx.ALL, 6)        
        self.m_isUseTwoLayers = wx.CheckBox(panel, label="Use 2\nLayers")        
        self.Bind(wx.EVT_CHECKBOX, self.OnIsUseTwoLayers) 
        hor_box_NN_parameters.Add(self.m_isUseTwoLayers, 0, wx.ALL, 6)        
        m_textL1 = wx.StaticText(panel, -1, "Neurons\nin Layer 1")
        hor_box_NN_parameters.Add(m_textL1, 0, wx.ALL, 6)
        self.m_neuronsInL1 = wx.TextCtrl(panel, -1, "50", size=(50,-1))
        hor_box_NN_parameters.Add(self.m_neuronsInL1, 0, wx.ALL, 6)
        m_textL2 = wx.StaticText(panel, -1, "Neurons\nin Layer 2")
        hor_box_NN_parameters.Add(m_textL2, 0, wx.ALL, 6)
        self.m_neuronsInL2 = wx.TextCtrl(panel, -1, "20", size=(50,-1))
        hor_box_NN_parameters.Add(self.m_neuronsInL2, 0, wx.ALL, 6)
        self.m_neuronsInL2.Enable(False)
        vert_box.Add(hor_box_NN_parameters, 0, wx.ALL, 6)
        
        m_textR1 = wx.StaticText(panel, -1, "Train and evaluate\nNeural Network.")
        hor_box_refresh.Add(m_textR1, 0, wx.ALL, 4)                
        self.m_refresh = wx.Button(panel, -1, "Evaluate")
        self.m_refresh.Bind(wx.EVT_BUTTON, self.OnRefresh)
        hor_box_refresh.Add(self.m_refresh, 0, wx.ALL, 6)
        self.m_evalResults = wx.TextCtrl(panel, -1, "(Evaluation results)", size=(1200,-1), style=wx.TE_READONLY)   # style=wx.TE_MULTILINE|wx.TE_READONLY
        hor_box_refresh.Add(self.m_evalResults, 0, wx.ALL, 10)        
        vert_box.Add(hor_box_refresh, 0, wx.ALL, 6)

        self.m_choice = wx.Choice(panel, choices = [], size=(1800,-1))
        self.m_choice.Bind(wx.EVT_CHOICE, self.OnChoice)
        vert_box.Add(self.m_choice, 0, wx.ALL, 10)

        self.m_searchText = wx.TextCtrl(panel, -1, "", size=(200,-1))
        hor_box_search.Add(self.m_searchText, 0, wx.ALL, 10)        
        m_textS1 = wx.StaticText(panel, -1, "Search for text\nin all data:")
        hor_box_search.Add(m_textS1, 0, wx.ALL, 6)
        self.m_search = wx.Button(panel, -1, "Search")
        self.m_search.Bind(wx.EVT_BUTTON, self.OnSearchText)
        hor_box_search.Add(self.m_search, 0, wx.ALL, 10)
        
        m_textC1 = wx.StaticText(panel, -1, "   ")
        hor_box_search.Add(m_textC1, 0, wx.ALL, 10)
        m_textC2 = wx.StaticText(panel, -1, "Current\nPage:")
        hor_box_search.Add(m_textC2, 0, wx.ALL, 6)
        self.m_openLink = wx.Button(panel, wx.ID_OPEN, "Open PDF")
        self.m_openLink.Bind(wx.EVT_BUTTON, self.OnOpenLink)
        hor_box_search.Add(self.m_openLink, 0, wx.ALL, 10)
        self.m_copyLink = wx.Button(panel, wx.ID_OPEN, "Copy Link")
        self.m_copyLink.Bind(wx.EVT_BUTTON, self.OnCopyLink)
        hor_box_search.Add(self.m_copyLink, 0, wx.ALL, 10)        
        self.m_openHtml = wx.Button(panel, wx.ID_OPEN, "Open Html")
        self.m_openHtml.Bind(wx.EVT_BUTTON, self.OnOpenHtml)
        hor_box_search.Add(self.m_openHtml, 0, wx.ALL, 10)
        self.m_openPng = wx.Button(panel, wx.ID_OPEN, "Refresh Png")
        self.m_openPng.Bind(wx.EVT_BUTTON, self.OnShowPng)
        hor_box_search.Add(self.m_openPng, 0, wx.ALL, 10)

        vert_box.Add(hor_box_search, 0, wx.ALL, 10)
        
        self.m_context = wx.TextCtrl(panel, -1, "Context = headers", size=(1800,-1), style=wx.TE_READONLY)   # style=wx.TE_MULTILINE|wx.TE_READONLY
        vert_box.Add(self.m_context, 0, wx.ALL, 10)
        
        self.m_extracted = wx.TextCtrl(panel, -1, "Extracted data", size=(1800,1500), style=wx.TE_MULTILINE|wx.TE_READONLY|wx.TE_RICH2)
        tcFont = self.m_extracted.GetFont()
        tcFont.SetFamily( wx.FONTFAMILY_TELETYPE )
        self.m_extracted.SetFont( tcFont )
        vert_box.Add(self.m_extracted, 0, wx.ALL, 10)
        
        self.m_bmPanel = wx.Panel(splitter)
        self.m_bitmap = wx.StaticBitmap(self.m_bmPanel, pos=(5, 5))
        
        # nltk.download("stopwords") # needed once only!

        self.m_choice.SetSelection(0)
        self.OnChoice(None)
        self.OnShowPng(None)
                
        panel.SetSizer(vert_box)
        
        splitter.SplitVertically(panel, self.m_bmPanel)
        splitter.SetMinimumPaneSize(400)
        
        panel.Layout()

        # Load whatever we have in the Results folder, i.e. results from the previous run...        
        self.LoadResults()
        
        print("Some data will be printed here when you evaluate...")


    def computeResults(self):
        global all_doc_pages_tvt
        global model
        global vocab
        global tokenizer

        all_doc_pages_tvt = []
        
        if os.path.isdir(resultsFolder1):
            for fl in os.listdir(resultsFolder1):
                os.remove(os.path.join(resultsFolder1, fl))
        else:
            os.makedirs(resultsFolder1)
            
        self.m_extracted.AppendText('   \n')
        self.m_extracted.AppendText('CREATING VOCABULARY\n')
        create_vocabulary(self)
        resp = wx.MessageBox('Vocabulary created', 'Computation proceeding', wx.OK | wx.ICON_INFORMATION)
        
        # Test vocab saving and reloading
        '''
        with open('vocab.pickle', 'wb') as fpoutp:
            pickle.dump(vocab, fpoutp)
        with open('vocab.pickle', 'rb') as fpinpp:
            vocab = pickle.load(fpinpp)
        '''
        
        '''
        self.m_extracted.AppendText('   \n')
        self.m_extracted.AppendText('LOADING TRAINING SET\n')
        load_training_set_data(self)
        resp = wx.MessageBox('Training set loaded', 'Computation proceeding', wx.OK | wx.ICON_INFORMATION)
        '''
        
        self.m_extracted.AppendText('   \n')
        self.m_extracted.AppendText('LOADING TRAINING AND VALIDATION SETS\n')
        create_x_train_x_test(self)
        resp = wx.MessageBox('Training and validation sets loaded', 'Computation proceeding', wx.OK | wx.ICON_INFORMATION)
        
        # Test tokenizer saving and reloading
        '''
        with open('token.pickle', 'wb') as fpoutt:
            pickle.dump(tokenizer, fpoutt)
        with open('token.pickle', 'rb') as fpinpt:
            tokenizer = pickle.load(fpinpt)
        '''
        
        self.m_extracted.AppendText('   \n')
        self.m_extracted.AppendText('COMPUTING Yx2; CREATING AND TRAINING KERAS MODEL\n')
        create_xy_train_valid(self)
        resp = wx.MessageBox('Keras model trained', 'Computation proceeding', wx.OK | wx.ICON_INFORMATION)
        
        # Test model saving and reloading; creates FOLDER 'keras_model'
        '''
        model.save('keras_model')
        model = load_model('keras_model')
        '''

        '''
        self.m_extracted.AppendText('   \n')
        self.m_extracted.AppendText('EVALUATING MODES\n')
        evaluate_modes(self)
        resp = wx.MessageBox('Vocabulary created', 'Computation proceeding', wx.OK | wx.ICON_INFORMATION)
        '''
        
        self.m_extracted.AppendText('   \n')
        self.m_extracted.AppendText('EVALUATING TEST SET\n')
        (nbr_true_pos, nbr_fals_pos, nbr_true_neg, nbr_fals_neg) = evaluate_test_set(self)

        resp = wx.MessageBox(  'nbr_true_pos: ' + str(nbr_true_pos) +
                             ', nbr_fals_pos: ' + str(nbr_fals_pos) +
                             ', nbr_true_neg: ' + str(nbr_true_neg) +
                             ', nbr_false_neg: ' + str(nbr_fals_neg) + '\n' +
                             'Precision: ' + ('%.3f' % (nbr_true_pos/(nbr_true_pos+nbr_fals_pos))) + 
                             ', Recall: ' + ('%.3f' % (nbr_true_pos/(nbr_true_pos+nbr_fals_neg))),
                             'Computation finished', wx.OK | wx.ICON_INFORMATION)        

        self.m_evalResults.SetValue('Evaluation results ' + ('(ALL smpl): ' if modelUseSamples==1 else '(Test set): ') +
                                    'nbr_true_pos: ' + str(nbr_true_pos) +
                                    ', nbr_false_pos: ' + str(nbr_fals_pos) +
                                    ', nbr_true_neg: ' + str(nbr_true_neg) +
                                    ', nbr_false_neg: ' + str(nbr_fals_neg) + '.   ' +
                                    'Precision: ' + ('%.3f' % (nbr_true_pos/(nbr_true_pos+nbr_fals_pos))) + 
                                    ', Recall: ' + ('%.3f' % (nbr_true_pos/(nbr_true_pos+nbr_fals_neg)))                                    )

    
    def OnRefresh(self, event):
        global modelUseSamples
        global isConvertToLowercase
        global isRemovePunctuation
        global handleNumbers_RmKpRp
        global isRemoveStopWords
        global minimumWordLength
        global minVocabularyOccur
        global scoreWordsMethod
        global duplicatePosSamples
        global isUseTwoNNLayers
        global neuronsInNNLayer1
        global neuronsInNNLayer2

        self.m_extracted.SetLabel("")
        self.m_extracted.Refresh()
        self.m_extracted.Update()
        
        # NN Evaluation options
        modelUseSamples = self.m_evaluateSamples.GetSelection()
        isConvertToLowercase = self.m_isLowercase.IsChecked()
        isRemovePunctuation = self.m_isRemovePuncts.IsChecked()
        handleNumbers_RmKpRp = self.m_processNumbers.GetSelection()
        isRemoveStopWords = self.m_isRemoveStopWs.IsChecked()
        minimumWordLength = int(self.m_shortLimit.GetValue())
        minVocabularyOccur = int(self.m_minOccurrence.GetValue())
        scoreWordsMethod = self.m_bagOfWordsMethod.GetString(self.m_bagOfWordsMethod.GetSelection())
        duplicatePosSamples = int(self.m_duplicatePos.GetValue())
        isUseTwoNNLayers = self.m_isUseTwoLayers.IsChecked()
        neuronsInNNLayer1 = int(self.m_neuronsInL1.GetValue())
        neuronsInNNLayer2 = int(self.m_neuronsInL2.GetValue())

        self.computeResults()
        
        '''
        truth_set = {tuple(x) for x in ground_truth_doc_pages}
        found_set = {tuple(x) for x in found_doc_pages}
        match_set = truth_set & found_set
        missd_set = truth_set - found_set
        bogus_set = found_set - truth_set
        precision = len(match_set)/(len(match_set)+len(bogus_set))
        recall    = len(match_set)/(len(match_set)+len(missd_set))
        resp = wx.MessageBox('Ground Truth items: ' + str(len(truth_set)) + '\nProgram found items: ' + str(len(found_set)) +
                             '\nMatch items: ' + str(len(match_set)) + ', missed items: ' + str(len(missd_set)) + ', bogus items: ' + str(len(bogus_set)) +
                             '\nPrecision: ' + ('%.3f' % precision) + ', Recall: ' + ('%.3f' % recall),
                             'Computation finished', wx.OK | wx.ICON_INFORMATION)
        '''

        if modelUseSamples==0:        
            all_pages = len(all_doc_pages_tvt)
            tra_pages = len([x for x in all_doc_pages_tvt if x["is_train"]])
            val_pages = len([x for x in all_doc_pages_tvt if x["is_valid"]])
            tst_pages = len([x for x in all_doc_pages_tvt if x["is_test" ]])
            resp = wx.MessageBox('All Pages: ' + str(all_pages) + '\nTrain pages: ' + str(tra_pages) +
                                 '\nValidation pages: ' + str(val_pages) + '\nTest pages: ' + str(tst_pages),
                                 'Pages divided into sets', wx.OK | wx.ICON_INFORMATION)
            
            all_gt_pgs = len([x for x in all_doc_pages_tvt if [x["doc"], x["page"]] in ground_truth_doc_pages])
            tra_gt_pgs = len([x for x in all_doc_pages_tvt if x["is_train"] and [x["doc"], x["page"]] in ground_truth_doc_pages])
            val_gt_pgs = len([x for x in all_doc_pages_tvt if x["is_valid"] and [x["doc"], x["page"]] in ground_truth_doc_pages])
            tst_gt_pgs = len([x for x in all_doc_pages_tvt if x["is_test" ] and [x["doc"], x["page"]] in ground_truth_doc_pages])
            resp = wx.MessageBox('All GT Pages: ' + str(all_gt_pgs) + '\nTrain GT pages: ' + str(tra_gt_pgs) +
                                 '\nValidation GT pages: ' + str(val_gt_pgs) + '\nTest GT pages: ' + str(tst_gt_pgs),
                                 'Pages divided into sets', wx.OK | wx.ICON_INFORMATION)

        '''
        # Verify that all Ground Truth pages are indeed found among all pages...
        for gt in ground_truth_doc_pages:
            for dp in all_doc_pages_tvt:
                if [dp["doc"], dp["page"]] == gt:
                    break
            else:
                resp = wx.MessageBox('Bad: Doc='+str(gt[0])+', Page='+str(gt[1]), 'Bad Item in G.T.', wx.OK | wx.ICON_WARNING)
        '''
        
        self.LoadResults()


    def LoadResults(self): 
        input_file = resultsFolder1 + "/NN_eval_results.txt"
        if not os.path.exists(input_file):
            return
        with open(input_file, "r") as file:
            self.m_data = file.readlines()

        self.m_links = []
        for d in self.m_data:
            if "file:///" in d:
                self.m_links.append(d)
                
        self.m_choice.Clear()
        self.m_choice.AppendItems(self.m_links)
        if self.m_choice.GetCount()>0:
            self.m_choice.SetSelection(0)
            self.OnChoice(None)

        

    def OnEvaluateSamples(self, event):
        pass


    def OnProcessNumbers(self, event):
        pass


    def OnBagOfWordsMethod(self, event):
        pass


    def OnIsRemoveShorts(self, event):
        pass


    def OnIsRemovePunctuations(self, event):
        pass


    def OnIsRemoveStopWords(self, event):
        pass


    def OnIsLowercase(self, event):
        pass


    def OnIsUseTwoLayers(self, event):
        self.m_neuronsInL2.Enable(self.m_isUseTwoLayers.IsChecked())
        pass


    def OnOpenLink(self, event):
        url = self.m_choice.GetString(self.m_choice.GetSelection())
        webbrowser.open(url)

    def OnCopyLink(self, event):
        url = self.m_choice.GetString(self.m_choice.GetSelection())
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(url))
            wx.TheClipboard.Close()

    def OnShowPng(self, event):
        if self.m_docPage:
            self.m_bmPanel.Refresh()
            imageFile = self.m_docPage
            imageFile = imageFile.replace(':', '')
            imageFile = imageFile.replace(' =>', '')
            imageFile = imageFile.replace(', ', '_')
            imageFile = imageFile.replace(' \n', '')
            imageFile = os.path.join(resultsFolder1, imageFile + '.png')
            png = wx.Image(imageFile, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            self.m_bitmap.SetBitmap(png) # size=(png.GetWidth(), png.GetHeight()))

    def OnOpenHtml(self, event):
        if self.m_docPage:
            htmlFile = self.m_docPage
            htmlFile = htmlFile.replace(':', '')
            htmlFile = htmlFile.replace(' =>', '')
            htmlFile = htmlFile.replace(', ', '_')
            htmlFile = htmlFile.replace(' \n', '')
            htmlFile = os.path.join(resultsFolder1, htmlFile + '.html')
            wx.LaunchDefaultBrowser(htmlFile)

    def OnSearchText(self, event):
        searchText = self.m_searchText.GetValue()
        if searchText:
            currentLink = self.m_choice.GetStringSelection();
            firstLink = ""
            firstText = ""
            lastLink = ""
            afterCurrent = False
            for d in self.m_data:
                if "file:///" in d:
                    lastLink = d
                if searchText in d:
                    if not afterCurrent:
                        if not firstLink:
                            firstLink = lastLink
                            firstText = d
                    elif lastLink:
                        self.m_choice.SetStringSelection(lastLink)
                        self.OnChoice(event)
                        '''
                        dlg = wx.MessageDialog(self, d,
                            "Text Found: "+searchText, wx.OK|wx.ICON_INFORMATION)
                        dlg.ShowModal()
                        dlg.Destroy()
                        '''
                        break
                if d==currentLink:
                    afterCurrent = True
                    lastLink = ""
            else:
                if firstLink:
                    self.m_choice.SetStringSelection(firstLink)
                    self.OnChoice(event)
                    dlg = wx.MessageDialog(self, firstText,
                        "Text Found: "+searchText + ' (from beginning)', wx.OK|wx.ICON_INFORMATION)
                    dlg.ShowModal()
                    dlg.Destroy()
                else:
                    dlg = wx.MessageDialog(self, 
                        "No such text in data",
                        searchText, wx.OK|wx.ICON_INFORMATION)
                    dlg.ShowModal()
                    dlg.Destroy()


    def OnClose(self, event):
        '''
        dlg = wx.MessageDialog(self, 
            "Do you really want to close this application?",
            "Confirm Exit", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
        '''
        self.Destroy()

    def OnAbout(self, event):
        dlg = AboutBox()
        dlg.ShowModal()
        dlg.Destroy()
  
    def OnChoice(self, event):
        idx = self.m_choice.GetSelection()
        if idx<0:
            return
        line = self.m_choice.GetString(idx);
        if not line:
            return
        foundStart = False
        foundTOC = False
        foundDoc = False
        extrText = ""
        self.m_docPage = ""
        for ln in self.m_data:
            if foundStart and foundTOC:
                if "file:///" in ln:
                    break
                else:
                    if not foundDoc:
                        self.m_docPage = ln
                        foundDoc = True
                    if ln.strip():
                        extrText += ln
            if foundStart and not foundTOC:
                self.m_context.SetLabel(ln)
                self.m_extracted.SetLabel('')
                foundTOC = True
            if line==ln:
                foundStart = True
        self.m_extracted.SetLabel(extrText)
        if self.m_docPage:
            self.OnShowPng(event)


app = wx.App(redirect=True)   # Error messages go to popup window
top = Frame("VesselAI_ExtractData_NN")
top.Show()
app.MainLoop()
