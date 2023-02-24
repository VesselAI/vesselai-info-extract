# -*- coding: utf-8 -*-
"""
Created on Mon May 30 10:10:10 2022

@author: TTERAK
"""
# The program tries to find the relevant data pages (containing e.g. fuel consumption info) in the *.PDFs listed in 'files_to_process'.
# This version uses some of the previously found pages (see program ExtractInfoPages_AllLines.py) as model pages (6 clusters); 
# if the word histogram of a candidate page matches that of any of the model page clusters, then that candidate page is accepted.
# When a relevant page is found, it is saved as a *.png and a *.html.

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
from collections import Counter
import re
import json


aboutText = """<p>Sorry, there is no information about this program. It is
running on version %(wxpy)s of <b>wxPython</b> and %(python)s of <b>Python</b>.
See <a href="http://wiki.wxpython.org">wxPython Wiki</a></p>""" 


resultsFolder1 = "Phase_1_Docs_Pages"
resultsFolder3 = "Phase_3_Docs_Pages"
    

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


nbr_clusters = 6 # Number of clusters; pages within a cluster are similar; this initial value is overridden in function computeResults()

# Model pages: Doc number, page number
learn_doc_pages = [
    [0, 79],
    [0, 86],
    [1, 93],
    [1, 78],
    [26, 59],
    [31, 33],
]

cluster_histograms = [Counter()] * nbr_clusters # this initial value is overridden in function computeResults()

cluster_word_counts = [0] * nbr_clusters # this initial value is overridden in function computeResults()

cluster_rel_lines = [
    { 'line_text': "Specific fuel oil consumption", 'line_offsets': [-4, -3, 0] },
    { 'line_text': "Specific fuel oil consumption", 'line_offsets': [-6, -5, 0, 1, 2] },
    { 'line_text': "Air flow rate", 'line_offsets': [0], 'line_text_2': "Engine output"},
    { 'line_text': "Spec. fuel consumption (g/kWh) with", 'line_offsets': [-5, -4, -3, -2, -1, 0, 1, 2] },
    { 'line_text': "Flow of air at 100% load", 'line_offsets': [-7, -6, -5, -4, -3, -2, -1, 0] },
    { 'line_text': "SFOC at 75% load (LFO)", 'line_offsets': [-5, -4, -3, -2, -1, 0, 1] }
    ]

# Maximum word count histogram difference between model page and candidate page, for the candidate page to be accepted
max_hist_diff = 0.5


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
    # In many MAN files, vertical text is marked with white text; let's eliminate those...
    white = ln.find("color:#ffffff")
    if white>0:
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


# Find out whether 'str' is a number (int or float) or an int with commas (as thousands separator) or <something>: number
def isNumber(str):
    str1 = str[:]
    # Find out whether the string contains a colon; if it does, only consider the text after the colon
    idx = str.find(': ')
    if idx>=0:
        str1 = str[idx+2:]
    else:
        idx = str.find(':')
        if idx>=0:
            str1 = str[idx+1:]

    # If there are commas or spaces as thousands separator, remove them
    str1 = str1.replace(',', '')
    str1 = str1.replace(' ', '')

    # And now finally: try to convert the string into a floating-point number; if this succeeds, return True; else False            
    try:
        float(str1)
        return True
    except ValueError:
        return False


def computeDifferences(nbrWords, counts):
    diff = [0.0] * nbr_clusters

    for c in range(nbr_clusters):
        for key in cluster_histograms[c].keys():
            if isNumber(key):
                continue
            if key in counts.keys():
                diff[c] += abs(counts[key]/nbrWords - cluster_histograms[c][key]/cluster_word_counts[c])
            else:
                diff[c] += cluster_histograms[c][key]/cluster_word_counts[c]
        for key in counts.keys():
            if isNumber(key):
                continue
            if not key in cluster_histograms[c].keys():
                diff[c] += counts[key]/nbrWords
            
    return diff


def findRelevantLines(linesToPrint, i_min_d):
    ret_lines = ""
    if i_min_d>=0 and i_min_d<nbr_clusters:
        line_text = ""
        line_offsets = [0]
        line_text_2 = ""
        line_offsets_2 = [0]
        line_text_3 = ""
        line_offsets_3 = [0]
        
        if 'line_text' in cluster_rel_lines[i_min_d]:
            line_text = cluster_rel_lines[i_min_d]['line_text']
        line_offsets = cluster_rel_lines[i_min_d]['line_offsets']
        if 'line_offsets' in cluster_rel_lines[i_min_d]:
            if len(line_offsets)<1:
                line_offsets = [0]
        if 'line_text_2' in cluster_rel_lines[i_min_d]:
            line_text_2 = cluster_rel_lines[i_min_d]['line_text_2']
        if 'line_offsets_2' in cluster_rel_lines[i_min_d]:
            line_offsets_2 = cluster_rel_lines[i_min_d]['line_offsets_2']
            if len(line_offsets_2)<1:
                line_offsets_2 = [0]
        if 'line_text_3' in cluster_rel_lines[i_min_d]:
            line_text_3 = cluster_rel_lines[i_min_d]['line_text_3']
        if 'line_offsets_3' in cluster_rel_lines[i_min_d]:
            line_offsets_3 = cluster_rel_lines[i_min_d]['line_offsets_3']
            if len(line_offsets_3)<1:
                line_offsets_3 = [0]
                
        for idx, ln in enumerate(linesToPrint):
            if line_text!="" and line_text in ln:
                for i in line_offsets:
                    if idx+i>=0 and idx+i<len(linesToPrint):
                        ret_lines += linesToPrint[idx+i] + "\n"
            if line_text_2!="" and line_text_2 in ln:
                for i in line_offsets_2:
                    if idx+i>=0 and idx+i<len(linesToPrint):
                        ret_lines += linesToPrint[idx+i] + "\n"
            if line_text_3!="" and line_text_3 in ln:
                for i in line_offsets_3:
                    if idx+i>=0 and idx+i<len(linesToPrint):
                        ret_lines += linesToPrint[idx+i] + "\n"

    if len(ret_lines)<=0:
        ret_lines = "Cluster "+str(i_min_d)+", key text: "+line_text
    return ret_lines        
        

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


def printTocItems(doc, pgNum1, fp):
    # Occasionally headers contain non-printing characters; print what can be printed, skip the rest
    try:
        print(findTocItems(doc, pgNum1), file=fp)
    except:
        toc_list = findTocItems(doc, pgNum1)
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



class HtmlWindow(wx.html.HtmlWindow):
    def __init__(self, parent, id, size=(600,400)):
        wx.html.HtmlWindow.__init__(self,parent, id, size=size)
        if "gtk2" in wx.PlatformInfo:
            self.SetStandardFonts()

    def OnLinkClicked(self, link):
        wx.LaunchDefaultBrowser(link.GetHref())

        
class AboutBox(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, None, -1, "About VesselAI_ExtractData_Phase_3",
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

        splitter = wx.SplitterWindow(self, style=wx.SP_NO_XP_THEME | wx.SP_3D | wx.SP_LIVE_UPDATE)
        panel = wx.Panel(splitter)
        vert_box = wx.BoxSizer(wx.VERTICAL)
        hor_box_modelPs = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_sLines0 = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_sLines1 = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_refresh = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_exclude = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_search = wx.BoxSizer(wx.HORIZONTAL)
        
        self.m_docPage = ""
        # file:///<file_path>.pdf#page=<page_number> entries will be added to this array for those pages where 'Exclude This Page from Results' check box is set.
        self.m_excludedPages = [] 
        input_file_3 = resultsFolder3 + "/Phase3_excluded.txt"
        if os.path.exists(input_file_3):
            with open(input_file_3, 'r') as fp2:
                self.m_excludedPages = fp2.readlines()

        m_textP2 = wx.StaticText(panel, -1, "\nData pages from phase 1 are in this upper drop-down. Cluster center pages from phase 2 are in 'Model Pages' list. You can Add new model pages from the drop-down to the list, or Delete existing ones.")
        vert_box.Add(m_textP2, 0, wx.ALL, 6)
        
        input_file_1 = resultsFolder1 + "/Phase1_results.txt"
        with open(input_file_1) as fp1:
            prev_data = fp1.readlines()
        input_file_e = resultsFolder1 + "/Phase1_excluded.txt"
        with open(input_file_e) as fpe:
            excluded = fpe.readlines()
        self.m_data1 = []
        for d in prev_data:
            if d not in excluded:
                self.m_data1.append(d)

        self.m_links1 = []
        for d in self.m_data1:
            if "file:///" in d:
                self.m_links1.append(d)
        self.m_prevChoice = wx.Choice(panel, choices = self.m_links1, size=(1800,-1))
        self.m_prevChoice.Bind(wx.EVT_CHOICE, self.OnPrevChoice)
        vert_box.Add(self.m_prevChoice, 0, wx.ALL, 2)
        
        vert_box_MP = wx.BoxSizer(wx.VERTICAL)
        vert_box_MP_Name = wx.BoxSizer(wx.VERTICAL)
        vert_box_MP_Line = wx.BoxSizer(wx.VERTICAL)
        vert_box_MP_Min = wx.BoxSizer(wx.HORIZONTAL)
        vert_box_MP_Max = wx.BoxSizer(wx.HORIZONTAL)
        m_textM1 = wx.StaticText(panel, -1, "Model\nPages:")
        vert_box_MP.Add(m_textM1, 0, wx.ALL, 4)
        self.m_addMP = wx.Button(panel, -1, "Add/Mod")
        self.m_addMP.Bind(wx.EVT_BUTTON, self.OnAddModelPage)
        vert_box_MP.Add(self.m_addMP, 0, wx.ALL, 4)
        self.m_delMP = wx.Button(panel, -1, "Delete")
        self.m_delMP.Bind(wx.EVT_BUTTON, self.OnDelModelPage)
        vert_box_MP.Add(self.m_delMP, 0, wx.ALL, 4)
        hor_box_modelPs.Add(vert_box_MP, 0, wx.ALL, 10)
        
        # These default values are overwritten with values from file 'Phase3_modelPages.json', if this file exists
        self.m_modelPages = [
            { "name": "Fuel oil 1", "doc_index": 0, "page_nbr": 79, 
             "line_1_text": "Specific fuel oil consumption", "line_1_offsets": [-4, -3, 0], 
             "line_2_text": "", "line_2_offsets": [], "line_3_text": "", "line_3_offsets": [] },
            
            { "name": "Fuel oil 2", "doc_index": 0, "page_nbr": 86, 
             "line_1_text": "Specific fuel oil consumption", "line_1_offsets": [-6, -5, 0, 1, 2], 
             "line_2_text": "", "line_2_offsets": [], "line_3_text": "", "line_3_offsets": [] },
            
            { "name": "Air flow 3", "doc_index": 1, "page_nbr": 93, 
             "line_1_text": "Air flow rate", "line_1_offsets": [0], 
             "line_2_text": "Engine output", "line_2_offsets": [0], "line_3_text": "", "line_3_offsets": [] },
            
            { "name": "Fuel oil 4", "doc_index": 1, "page_nbr": 78, 
             "line_1_text": "Spec. fuel consumption (g/kWh) with", "line_1_offsets": [-5, -4, -3, -2, -1, 0, 1, 2], 
             "line_2_text": "", "line_2_offsets": [], "line_3_text": "", "line_3_offsets": [] },
            
            { "name": "Air flow 5", "doc_index": 26, "page_nbr": 59, 
             "line_1_text": "Flow of air at 100% load", "line_1_offsets": [-7, -6, -5, -4, -3, -2, -1, 0], 
             "line_2_text": "", "line_2_offsets": [], "line_3_text": "", "line_3_offsets": [] },
            
            { "name": "SFOC 6", "doc_index": 31, "page_nbr": 33, 
             "line_1_text": "SFOC at 75% load (LFO)", "line_1_offsets": [-5, -4, -3, -2, -1, 0, 1], 
             "line_2_text": "", "line_2_offsets": [], "line_3_text": "", "line_3_offsets": [] },
            ]
        
        input_file_j = resultsFolder3 + "/Phase3_modelPages.json"
        try:
            fileJson = open(input_file_j, "r")
            jsonContent = fileJson.read()
            loadedModelPages = json.loads(jsonContent)
            self.m_modelPages = []
            for mp in loadedModelPages:
                if not mp["excluded"]:
                    mpe = { "name": mp["name"], "doc_index": mp["doc_index"], "page_nbr": mp["page_nbr"] }
                    self.m_modelPages.append(mpe)
                
            for mp in self.m_modelPages:
                if not "line_1_text" in mp:
                    mp["line_1_text"] = ""
                if not "line_1_offsets" in mp:
                    mp["line_1_offsets"] = []
                if not "line_2_text" in mp:
                    mp["line_2_text"] = ""
                if not "line_2_offsets" in mp:
                    mp["line_2_offsets"] = []
                if not "line_3_text" in mp:
                    mp["line_3_text"] = ""
                if not "line_3_offsets" in mp:
                    mp["line_3_offsets"] = []
        except IOError:
            pass # using the default self.m_modelPages data defined above...

        mp_names = [i["name"] for i in self.m_modelPages]
        self.m_modelPagesList = wx.ListBox(panel, choices = mp_names, size=(150,110))
        self.m_modelPagesList.Bind(wx.EVT_LISTBOX, self.OnSelectModelPage)
        hor_box_modelPs.Add(self.m_modelPagesList, 0, wx.ALL, 10)
        
        m_textM2 = wx.StaticText(panel, -1, "Selected\nModel\nPage:")
        hor_box_modelPs.Add(m_textM2, 0, wx.ALL, 16)
        m_textM3 = wx.StaticText(panel, -1, "Name")
        vert_box_MP_Name.Add(m_textM3, 0, wx.ALL, 6)
        self.m_modelPageName = wx.TextCtrl(panel, -1, "(name)", size=(170,-1))
        vert_box_MP_Name.Add(self.m_modelPageName, 0, wx.ALL, 2)        
        self.m_modelPageDoc = wx.TextCtrl(panel, -1, "(doc, pg)", size=(170,-1), style=wx.TE_READONLY)
        vert_box_MP_Name.Add(self.m_modelPageDoc, 0, wx.ALL, 2)        
        hor_box_modelPs.Add(vert_box_MP_Name, 0, wx.ALL, 10)
        m_textM6 = wx.StaticText(panel, -1, "Model Pages are pages that contain data relevant to you.\nAfter you have defined a number of Model Pages, "+
                                 "the program tries\nto find more of similar pages, which appear in the lower drop-down.\n"+
                                 "\nTo rename of an existing Model Page, select Model Page in the list, enter new name for it, and press 'Rename'.\n"+
                                 "To add a new Model Page, select a new page in the upper drop-down, enter Name for it, and press 'Add'.")
        hor_box_modelPs.Add(m_textM6, 0, wx.ALL, 10)
        
        vert_box.Add(hor_box_modelPs, 0, wx.ALL, 10)

        m_textL1 = wx.StaticText(panel, -1, "Lines\nto Save:   ")
        hor_box_sLines0.Add(m_textL1, 0, wx.ALL, 4)
        m_textL2 = wx.StaticText(panel, -1, "For each type of Model page, usually only a few lines contain the data we want to save."+
                                 " We recognize such lines by a text snippet in that line (usually text in header column).\n"+
                                 "Besides such line itself, we can save some lines before and after it, by entering Offsets to line index:"+
                                 " -1 = previous line, 1 = next line. Separate offsets with spaces. Max 3 recognized lines.")
        hor_box_sLines0.Add(m_textL2, 0, wx.ALL, 4)
        vert_box.Add(hor_box_sLines0, 0, wx.ALL, 2)

        m_textL3 = wx.StaticText(panel, -1, "Text in\nLine 1")
        hor_box_sLines1.Add(m_textL3, 0, wx.ALL, 4)
        self.m_textInLine1 = wx.TextCtrl(panel, -1, "(text 1)", size=(160,-1))
        hor_box_sLines1.Add(self.m_textInLine1, 0, wx.ALL, 4)        
        m_textL4 = wx.StaticText(panel, -1, "Offsets to\nLine 1")
        hor_box_sLines1.Add(m_textL4, 0, wx.ALL, 4)
        self.m_offsetsList1 = wx.TextCtrl(panel, -1, "(offsets 1)", size=(150,-1))
        hor_box_sLines1.Add(self.m_offsetsList1, 0, wx.ALL, 4)        
        
        m_textL5 = wx.StaticText(panel, -1, "      Line 2:\n      Text")
        hor_box_sLines1.Add(m_textL5, 0, wx.ALL, 4)
        self.m_textInLine2 = wx.TextCtrl(panel, -1, "(text 2)", size=(100,-1))
        hor_box_sLines1.Add(self.m_textInLine2, 0, wx.ALL, 4)        
        m_textL6 = wx.StaticText(panel, -1, "Line2:\nOffsets")
        hor_box_sLines1.Add(m_textL6, 0, wx.ALL, 4)
        self.m_offsetsList2 = wx.TextCtrl(panel, -1, "(offsets 2)", size=(100,-1))
        hor_box_sLines1.Add(self.m_offsetsList2, 0, wx.ALL, 4)        
        
        m_textL7 = wx.StaticText(panel, -1, "      Line 3:\n      Text")
        hor_box_sLines1.Add(m_textL7, 0, wx.ALL, 4)
        self.m_textInLine3 = wx.TextCtrl(panel, -1, "(text 3)", size=(80,-1))
        hor_box_sLines1.Add(self.m_textInLine3, 0, wx.ALL, 4)        
        m_textL8 = wx.StaticText(panel, -1, "Line 3:\nOffsets")
        hor_box_sLines1.Add(m_textL8, 0, wx.ALL, 4)
        self.m_offsetsList3 = wx.TextCtrl(panel, -1, "(offsets 3)", size=(80,-1))
        hor_box_sLines1.Add(self.m_offsetsList3, 0, wx.ALL, 4)        

        vert_box.Add(hor_box_sLines1, 0, wx.ALL, 2)


        m_textR1 = wx.StaticText(panel, -1, "After you have added enough pages to 'Model Pages' above, press button 'Refresh', " + 
                                "and the drop-down below is filled\nwith matching data pages. " +
                                "When you have checked that the desired data pages are indeed in the drop-down,\npress button 'Results', " +
                                "and a window showing key results is opened. Parameter 'Max Hist Diff' affects page matches.")
        hor_box_refresh.Add(m_textR1, 0, wx.ALL, 10)
        m_textR2 = wx.StaticText(panel, -1, "Max Hist Diff\n(>0, <~1)\nsmall: less pages")
        hor_box_refresh.Add(m_textR2, 0, wx.ALL, 10)
        self.m_maxHistDiff = wx.TextCtrl(panel, -1, "0.5", size=(50,-1))
        hor_box_refresh.Add(self.m_maxHistDiff, 0, wx.ALL, 10)        
        self.m_refresh = wx.Button(panel, -1, "Refresh")
        self.m_refresh.Bind(wx.EVT_BUTTON, self.OnRefresh)
        hor_box_refresh.Add(self.m_refresh, 0, wx.ALL, 10)
        self.m_nextPhase = wx.Button(panel, -1, "Results")
        self.m_nextPhase.Bind(wx.EVT_BUTTON, self.OnNextPhase)
        hor_box_refresh.Add(self.m_nextPhase, 0, wx.ALL, 10)
        vert_box.Add(hor_box_refresh, 0, wx.ALL, 6)

        m_textE1 = wx.StaticText(panel, -1, "You may wish to eliminate some pages from the results. Select result page from the drop-down below, and check\n'Exclude This Page...'. "+
                                 "If you intend to exclude most of the result pages, button 'Exclude All' checks it for all pages.")
        hor_box_exclude.Add(m_textE1, 0, wx.ALL, 10)        
        self.m_excludePage = wx.CheckBox(panel, label="Exclude This Page\nfrom Results")        
        self.Bind(wx.EVT_CHECKBOX, self.OnExcludePage) 
        hor_box_exclude.Add(self.m_excludePage, 0, wx.ALL, 6)
        self.m_excludeAll = wx.Button(panel, wx.ID_OPEN, "Exclude All")
        self.m_excludeAll.Bind(wx.EVT_BUTTON, self.OnExcludeAll)
        hor_box_exclude.Add(self.m_excludeAll, 0, wx.ALL, 10)        
        self.m_includeAll = wx.Button(panel, wx.ID_OPEN, "Include All")
        self.m_includeAll.Bind(wx.EVT_BUTTON, self.OnIncludeAll)
        hor_box_exclude.Add(self.m_includeAll, 0, wx.ALL, 10)
        vert_box.Add(hor_box_exclude, 0, wx.ALL, 6)

        self.m_choice = wx.Choice(panel, choices = [], size=(1800,-1))
        self.m_choice.Bind(wx.EVT_CHOICE, self.OnChoice)
        vert_box.Add(self.m_choice, 0, wx.ALL, 2)

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
        
        self.m_ourData = wx.TextCtrl(panel, -1, "Selected data, if any", size=(1800,100), style=wx.TE_MULTILINE|wx.TE_READONLY|wx.TE_RICH2)
        tcFont = self.m_ourData.GetFont()
        tcFont.SetFamily( wx.FONTFAMILY_TELETYPE )
        self.m_ourData.SetFont( tcFont )
        vert_box.Add(self.m_ourData, 0, wx.ALL, 10)
        
        self.m_extracted = wx.TextCtrl(panel, -1, "Extracted data", size=(1800,1500), style=wx.TE_MULTILINE|wx.TE_READONLY|wx.TE_RICH2)
        tcFont = self.m_extracted.GetFont()
        tcFont.SetFamily( wx.FONTFAMILY_TELETYPE )
        self.m_extracted.SetFont( tcFont )
        vert_box.Add(self.m_extracted, 0, wx.ALL, 10)
        
        self.m_bmPanel = wx.Panel(splitter)
        self.m_bitmap = wx.StaticBitmap(self.m_bmPanel, pos=(5, 5))
        
        self.m_choice.SetSelection(0)
        self.OnChoice(None)
        self.OnShowPng(None)
                
        panel.SetSizer(vert_box)
        
        splitter.SplitVertically(panel, self.m_bmPanel)
        splitter.SetMinimumPaneSize(600)

        panel.Layout()
                
        # Load whatever we have in the Results folder, i.e. results from the previous run...        
        self.LoadResults()

        
    def computeResults(self):
        global nbr_clusters
        global learn_doc_pages
        global cluster_histograms
        global cluster_word_counts
        global cluster_rel_lines
        global max_hist_diff
        nbr_clusters = len(self.m_modelPages)
        learn_doc_pages = [[i["doc_index"],i["page_nbr"]] for i in self.m_modelPages]        
        cluster_histograms = [Counter()] * nbr_clusters        
        cluster_word_counts = [0] * nbr_clusters        
        cluster_rel_lines = [{ 'line_text': i['line_1_text'], 'line_offsets': i['line_1_offsets'], 
                              'line_text_2': i['line_2_text'], 'line_offsets_2': i['line_2_offsets'], 
                              'line_text_3': i['line_3_text'], 'line_offsets_3': i['line_3_offsets'] } for i in self.m_modelPages]
        max_hist_diff = float(self.m_maxHistDiff.GetValue())
    
        if os.path.isdir(resultsFolder3):
            for i in os.listdir(resultsFolder3):
                os.remove(os.path.join(resultsFolder3, i))
        else:
            os.makedirs(resultsFolder3)
            
        output_file = resultsFolder3 + "/Phase3_results.txt"
        outp_fl_key = resultsFolder3 + "/Phase3_all_key_results.txt"
        with open(output_file, 'w') as fp:
            with open(outp_fl_key, 'w') as fpkey:
                
                self.m_extracted.AppendText('Learning...\n')
                # First go through the files; if 'docNumber' does not appear in 'learn_doc_pages' etc., skip the file
                for docNumber, filename in enumerate(files_to_process):
                
                    found = False
                    for c in learn_doc_pages:
                        if c[0]==docNumber:
                            found = True
                            break
                
                    if found:            
                        # creating a pdf file object 
                        doc = fitz.open(filename)
                        self.m_extracted.AppendText(filename+'\n')
                        self.m_extracted.Refresh()
                        self.m_extracted.Update()
                        # creating a page object
                        for c_idx, c in enumerate(learn_doc_pages):
                            for pageNumber, page in enumerate(doc):
                                if c[0]==docNumber and c[1]==pageNumber+1:
                                    pageWords = page.get_text('words')
                                    wordsOnly = [pg[4] for pg in pageWords]
                                    
                                    nbrWords = len(wordsOnly)
                                    counts = Counter(wordsOnly)
                                        
                                    # Compiler / library bug, work-around...
                                    prevCount = cluster_histograms[c_idx] 
                                    prevCount = prevCount + counts
                                    cluster_histograms[c_idx] = prevCount
                                    cluster_word_counts[c_idx] += nbrWords
                      
                self.m_extracted.AppendText('\nProcessing...\n')
                # Now go through the files again; for each page check whether the histogram is similar to those of fuel_consum etc. 
                for docNumber, filename in enumerate(files_to_process):
                
                    # creating a pdf file object 
                    doc = fitz.open(filename)
                      
                    self.m_extracted.AppendText(filename+'\n')
                    self.m_extracted.Refresh()
                    self.m_extracted.Update()
                    
                    prev_lines = []
                    # creating a page object
                    for pageNumber, page in enumerate(doc):
                        # extracting text from page
                        pageWords = page.get_text('words')
                        wordsOnly = [pg[4] for pg in pageWords]
                        nbrWords = len(wordsOnly)
                        counts = Counter(wordsOnly)
                        # Value 1.0 creates a first hit at Doc 0, Page 79 (=model page). The diffs are: 0.1666, 0.5371, 1.094, 1.118, 1.368, 1.458
                        # Value 1.0 finds a few irrelevant pages, but value 0.2 finds the model pages only
                        # So we came down from 1.0 (testing results of values 0.8 and 0.6) until the irrelevant pages were eliminated! 
                        # Some bad pages still at 0.5, but at 0.4 some good ones were eliminated, too (and still some bad pages!)...
                        diffs = computeDifferences(nbrWords, counts)
                        if min(diffs) <= max_hist_diff:
                            i_min_d = diffs.index(min(diffs))
                            pix = page.get_pixmap()
                            png_filename = resultsFolder3 + "/Doc" + str(docNumber) + "_page" + str(pageNumber+1) + ".png"
                            pix.save(png_filename)
                            url = "file:///" + os.path.abspath(filename) + "#page=" + str(pageNumber+1)
                            print(url, file=fp)
                            print(url, file=fpkey)
                            
                            printTocItems(doc, pageNumber+1, fp)
                            printTocItems(doc, pageNumber+1, fpkey)
                            
                            # extracting html from page
                            pageHtml = page.get_text('html', sort=True)
                            html_filename = resultsFolder3 + "/Doc" + str(docNumber) + "_page" + str(pageNumber+1) + ".html"
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
                                      
                            self.m_extracted.AppendText("   Doc:"+str(docNumber)+", Page:"+str(pageNumber+1)+"\n")
                            print("Doc:"+str(docNumber)+", Page:"+str(pageNumber+1)+" => ", file=fp)
                            
                            ourData = findRelevantLines(linesToPrint, i_min_d)
                            print("{", file=fp)
                            print(ourData, file=fp)
                            print(ourData, file=fpkey)
                            print("}", file=fp)
                
                            for ltp in linesToPrint:
                                try:
                                    print(ltp.rstrip(), file=fp)
                                except:
                                    pass
                                
                allLogText = []
                nbrLogLines = self.m_extracted.GetNumberOfLines()
                for ll in range(nbrLogLines):
                    allLogText.append(self.m_extracted.GetLineText(ll))
                dlg = wx.MessageDialog(self, 
                    "Log data are available in memory.\nDo you want to Copy the data to Clipboard?",
                    "Copy Log?", wx.YES_NO|wx.NO_DEFAULT|wx.ICON_QUESTION)
                result = dlg.ShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    dataObj = wx.TextDataObject()
                    dataObj.SetText("\n".join(allLogText))
                    wx.TheClipboard.Open()
                    wx.TheClipboard.SetData(dataObj)
                    wx.TheClipboard.Close()


    def OnPreviousPhase(self, event):
        # Maybe we will have an entry for this in the File menu...
        pass

    
    def OnSaveProject(self, event):
        # Maybe we will have an entry for this in the File menu...
        pass

    
    def OnPrevChoice(self, event):
        line = self.m_prevChoice.GetString(self.m_prevChoice.GetSelection());
        if not line:
            return
        foundStart = False
        foundTOC = False
        foundDoc = False
        extrText = ""
        self.m_docPage = ""
        for ln in self.m_data1:
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
        self.m_ourData.SetLabel("")
        self.m_extracted.SetLabel(extrText)
        if self.m_docPage:
            self.OnShowPng(event, True)
            
            docu_idx = -1
            page_idx = -1
            int_list = re.findall('\d+', self.m_docPage)
            if len(int_list)==2:
                docu_idx = int(int_list[0])
                page_idx = int(int_list[1])
                for idx, mp in enumerate(self.m_modelPages):
                    if mp["doc_index"]==docu_idx and mp["page_nbr"]==page_idx:
                        self.m_modelPagesList.SetSelection(idx)
                        self.OnSelectModelPage(None)
                        break
                else:
                    self.m_modelPagesList.SetSelection(-1)
                    self.OnSelectModelPage(None)

    
    def OnRefresh(self, event):
        self.m_extracted.SetLabel("")
        self.m_extracted.Refresh()
        self.m_extracted.Update()

        self.computeResults()
        
        self.LoadResults()


    def LoadResults(self): 
        input_file = resultsFolder3 + "/Phase3_results.txt"
        if not os.path.exists(input_file):
            return
        with open(input_file, "r") as file:
            self.m_data3 = file.readlines()

        self.m_links3 = []
        for d in self.m_data3:
            if "file:///" in d:
                self.m_links3.append(d)
                
        self.m_choice.Clear()
        self.m_choice.AppendItems(self.m_links3)
        if self.m_choice.GetCount()>0:
            self.m_choice.SetSelection(0)
            self.OnChoice(None)

    
    def OnNextPhase(self, event):
        output_file = resultsFolder3 + "/Phase3_excluded.txt"
        with open(output_file, 'w') as fp:
            for pg in self.m_excludedPages:
                print(pg, end='', file=fp)
                
        input_file_3 = resultsFolder3 + "/Phase3_all_key_results.txt"
        with open(input_file_3) as fp2:
            all_key_data = fp2.readlines()
        input_file_e = resultsFolder3 + "/Phase3_excluded.txt"
        with open(input_file_e) as fpe:
            excluded = fpe.readlines()
        write_line = True
        output_file = resultsFolder3 + "/Phase3_key_results.txt"
        with open(output_file, "w") as fpo:
            for line in all_key_data:
                if "file:///" in line:
                    if line in excluded:
                        write_line = False
                    else:
                        write_line = True
                if write_line:
                    print(line, file=fpo, end='')

        output_file = resultsFolder3 + "/Phase3_modelPages.json"
        jsonString = json.dumps(self.m_modelPages)
        jsonFile = open(output_file, "w")
        jsonFile.write(jsonString)
        jsonFile.close()        

        os.system("notepad " + resultsFolder3 + "/Phase3_key_results.txt")


    def OnSelectModelPage(self, event):
        idx = self.m_modelPagesList.GetSelection()
        if idx>=0:
            self.m_modelPageName.SetValue(self.m_modelPages[idx]["name"])
            self.m_textInLine1.SetValue(self.m_modelPages[idx]["line_1_text"])
            str_offs_1 = [str(i) for i in self.m_modelPages[idx]["line_1_offsets"]]
            self.m_offsetsList1.SetValue(' '.join(str_offs_1))
            self.m_textInLine2.SetValue(self.m_modelPages[idx]["line_2_text"])
            str_offs_2 = [str(i) for i in self.m_modelPages[idx]["line_2_offsets"]]
            self.m_offsetsList2.SetValue(' '.join(str_offs_2))
            self.m_textInLine3.SetValue(self.m_modelPages[idx]["line_3_text"])
            str_offs_3 = [str(i) for i in self.m_modelPages[idx]["line_3_offsets"]]
            self.m_offsetsList3.SetValue(' '.join(str_offs_3))
            doc_index = self.m_modelPages[idx]["doc_index"]
            page_nbr = self.m_modelPages[idx]["page_nbr"]
            str_doc_pg = "Doc " + str(doc_index) + ", Pg " + str(page_nbr)
            self.m_modelPageDoc.SetValue(str_doc_pg)
            self.m_addMP.SetLabel("Rename")
            if event!=None:
                url = "file:///" + os.path.abspath(files_to_process[doc_index]) + "#page=" + str(page_nbr) + "\n"
                c_idx = self.m_prevChoice.FindString(url)
                if c_idx>=0:
                    self.m_prevChoice.SetSelection(c_idx)
                    self.OnPrevChoice(None)
                else:
                    resp = wx.MessageBox(url, 'No such item in list', wx.OK | wx.ICON_WARNING)
        else:
            self.m_modelPageName.SetValue("")
            self.m_modelPageDoc.SetValue("(Enter Name, press Add)")
            self.m_textInLine1.SetValue("")
            self.m_offsetsList1.SetValue("")
            self.m_textInLine2.SetValue("")
            self.m_offsetsList2.SetValue("")
            self.m_textInLine3.SetValue("")
            self.m_offsetsList3.SetValue("")
            self.m_addMP.SetLabel("Add")


    def OnAddModelPage(self, event):
        name = self.m_modelPageName.GetValue()
        if not name.strip():
            resp = wx.MessageBox('No Name', 'Operation cancelled', wx.OK | wx.ICON_WARNING)
            return            
        if not self.m_docPage:
            resp = wx.MessageBox('No Doc, Page', 'Operation cancelled', wx.OK | wx.ICON_WARNING)
            return            

        txt1 = self.m_textInLine1.GetValue()
        off1 = self.m_offsetsList1.GetValue().split()
        map1 = map(int, off1)
        ofs1 = list(map1)
        txt2 = self.m_textInLine2.GetValue()
        off2 = self.m_offsetsList2.GetValue().split()
        map2 = map(int, off2)
        ofs2 = list(map2)
        txt3 = self.m_textInLine3.GetValue()
        off3 = self.m_offsetsList3.GetValue().split()
        map3 = map(int, off3)
        ofs3 = list(map3)
        
        docu_idx = -1
        page_idx = -1
        int_list = re.findall('\d+', self.m_docPage)
        if len(int_list)==2:
            docu_idx = int(int_list[0])
            page_idx = int(int_list[1])
            
        idx = self.m_modelPagesList.GetSelection()
        if idx>=0:
            self.m_modelPages[idx]["name"] = name.strip()
            self.m_modelPages[idx]["doc_index"] = docu_idx
            self.m_modelPages[idx]["page_nbr"] = page_idx
            self.m_modelPages[idx]["line_1_text"] = txt1
            self.m_modelPages[idx]["line_1_offsets"] = ofs1
            self.m_modelPages[idx]["line_2_text"] = txt2
            self.m_modelPages[idx]["line_2_offsets"] = ofs2
            self.m_modelPages[idx]["line_3_text"] = txt3
            self.m_modelPages[idx]["line_3_offsets"] = ofs3
        else:
            newitem = { "name": name, "doc_index": docu_idx, "page_nbr": page_idx, 
             "line_1_text": txt1, "line_1_offsets": ofs1, 
             "line_2_text": txt2, "line_2_offsets": ofs2, "line_3_text": txt3, "line_3_offsets": ofs3 }
            self.m_modelPages.append(newitem)
            
        mp_names = [i["name"] for i in self.m_modelPages]
        self.m_modelPagesList.Clear()
        self.m_modelPagesList.AppendItems(mp_names)


    def OnDelModelPage(self, event):
        idx = self.m_modelPagesList.GetSelection()
        if idx<0:
            resp = wx.MessageBox('No list item selected', 'Operation cancelled', wx.OK | wx.ICON_WARNING)
            return            
        self.m_modelPages.pop(idx)
        mp_names = [i["name"] for i in self.m_modelPages]
        self.m_modelPagesList.Clear()
        self.m_modelPagesList.AppendItems(mp_names)
        self.m_modelPageName.SetValue("")
        self.m_modelPageDoc.SetValue("")

    
    def OnExcludePage(self, event):
        fpage = self.m_choice.GetString(self.m_choice.GetSelection());
        if not fpage:
            return
        check = self.m_excludePage.IsChecked()
        if check:
            if not fpage in self.m_excludedPages:
                self.m_excludedPages.append(fpage)
        else:
            if fpage in self.m_excludedPages:
                self.m_excludedPages.remove(fpage)


    def OnExcludeAll(self, event):
        self.m_excludedPages = []
        for i in range(self.m_choice.GetCount()):
            self.m_excludedPages.append(self.m_choice.GetString(i))
        self.m_excludePage.SetValue(True)

    
    def OnIncludeAll(self, event):
        self.m_excludedPages = []
        self.m_excludePage.SetValue(False)
    
    
    def OnOpenLink(self, event):
        url = self.m_choice.GetString(self.m_choice.GetSelection())
        webbrowser.open(url)

    def OnCopyLink(self, event):
        url = self.m_choice.GetString(self.m_choice.GetSelection())
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(url))
            wx.TheClipboard.Close()

    def OnShowPng(self, event, folder1=False):
        if self.m_docPage:
            self.m_bmPanel.Refresh()
            imageFile = self.m_docPage
            imageFile = imageFile.replace(':', '')
            imageFile = imageFile.replace(' =>', '')
            imageFile = imageFile.replace(', ', '_')
            imageFile = imageFile.replace(' \n', '')
            if folder1:
                imageFile = os.path.join(resultsFolder1, imageFile + '.png')
            else:
                imageFile = os.path.join(resultsFolder3, imageFile + '.png')
            png = wx.Image(imageFile, wx.BITMAP_TYPE_ANY).ConvertToBitmap()
            self.m_bitmap.SetBitmap(png) # size=(png.GetWidth(), png.GetHeight()))

    def OnOpenHtml(self, event):
        if self.m_docPage:
            htmlFile = self.m_docPage
            htmlFile = htmlFile.replace(':', '')
            htmlFile = htmlFile.replace(' =>', '')
            htmlFile = htmlFile.replace(', ', '_')
            htmlFile = htmlFile.replace(' \n', '')
            htmlFile = os.path.join(resultsFolder3, htmlFile + '.html')
            wx.LaunchDefaultBrowser(htmlFile)


    def OnSearchText(self, event):
        searchText = self.m_searchText.GetValue()
        if searchText:
            currentLink = self.m_choice.GetStringSelection();
            firstLink = ""
            firstText = ""
            lastLink = ""
            afterCurrent = False
            for d in self.m_data3:
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
        for ln in self.m_data3:
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
        self.m_ourData.SetLabel("")
        findOpeningBrace = -1
        findClosingBrace = -1
        try:
            findOpeningBrace = extrText.index("{\n")
        except:
            pass
        try:
            findClosingBrace = extrText.index("}\n")
        except:
            pass
        if findOpeningBrace>=0 and findClosingBrace>=0 and findClosingBrace>findOpeningBrace:
            self.m_ourData.SetLabel(extrText[findOpeningBrace+2:findClosingBrace])
            extrText = extrText[0:findOpeningBrace] + extrText[findClosingBrace+1:]
        self.m_extracted.SetLabel(extrText)
        if self.m_docPage:
            self.OnShowPng(event)
        if line in self.m_excludedPages:
            self.m_excludePage.SetValue(True)
        else:
            self.m_excludePage.SetValue(False)


app = wx.App(redirect=True)   # Error messages go to popup window
top = Frame("VesselAI_ExtractData_Phase_3_ModelPages")
top.Show()
app.MainLoop()