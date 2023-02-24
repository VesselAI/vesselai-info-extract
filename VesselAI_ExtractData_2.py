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
import io
import html
import unicodedata
import string
from collections import Counter
import re
import json
import numpy as np
from sklearn.cluster import DBSCAN
from sklearn.cluster import AffinityPropagation
from sklearn.cluster import SpectralClustering
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
from sklearn.decomposition import TruncatedSVD
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import Normalizer


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


def computeDifference(nbrWords1, counts1, nbrWords2, counts2):
    diff = 0.0

    for key in counts1.keys():
        if isNumber(key):
            continue
        if key in counts2.keys():
            diff += abs(counts1[key]/nbrWords1 - counts2[key]/nbrWords2)
        else:
            diff += counts1[key]/nbrWords1
    for key in counts2.keys():
        if isNumber(key):
            continue
        if not key in counts1.keys():
            diff += counts2[key]/nbrWords2            
    return diff


class HtmlWindow(wx.html.HtmlWindow):
    def __init__(self, parent, id, size=(600,400)):
        wx.html.HtmlWindow.__init__(self,parent, id, size=size)
        if "gtk2" in wx.PlatformInfo:
            self.SetStandardFonts()

    def OnLinkClicked(self, link):
        wx.LaunchDefaultBrowser(link.GetHref())

        
class AboutBox(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, None, -1, "About VesselAI_ExtractData_Phase_2",
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
        hor_box_preComp = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_execute = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_isModel = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_search = wx.BoxSizer(wx.HORIZONTAL)
        
        self.m_docPage = ""
        
        m_textP2 = wx.StaticText(panel, -1, "\nData pages from previous phase are in this upper drop-down. " +
                                 "First select a clustering algorithm, Execute it and see how data pages are divided into clusters.")
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
        
        self.m_dataset_data = []
        self.m_dist_matrix = None
        
        m_textP3 = wx.StaticText(panel, -1, "Before you can try Clustering algorithms, you must load data and precompute Distance Matrix (takes a long time!).\n"+
                                 "After that the available Clustering algorithms appear in the drop-down below. Select one, enter parameter, and press Execute.")
        hor_box_preComp.Add(m_textP3, 0, wx.ALL, 10)
        self.m_precomp = wx.Button(panel, -1, "Precomp")
        self.m_precomp.Bind(wx.EVT_BUTTON, self.OnPrecompute)
        hor_box_preComp.Add(self.m_precomp, 0, wx.ALL, 10)
        vert_box.Add(hor_box_preComp, 0, wx.ALL, 6)
        
        m_textR1 = wx.StaticText(panel, -1, "The generated Data Page clusters appear in the list below.\n" + 
                                "When you select an item in the list (a cluster),\n" +
                                "the Data Pages belonging to that cluster appear in the lower drop-down.")
        hor_box_execute.Add(m_textR1, 0, wx.ALL, 10)
        self.m_clustAlgo = wx.Choice(panel, choices = ['(No distance matrix yet)'], size=(180,-1))
        self.m_clustAlgo.Bind(wx.EVT_CHOICE, self.OnClusterAlgo)
        self.m_clustAlgo.SetSelection(0)
        hor_box_execute.Add(self.m_clustAlgo, 0, wx.ALL, 10)
        self.m_textClustParam = wx.StaticText(panel, -1, "Cluster Param\n(first press\nPrecomp)")
        hor_box_execute.Add(self.m_textClustParam, 0, wx.ALL, 10)
        self.m_clusterParam = wx.TextCtrl(panel, -1, "1", size=(80,-1))
        hor_box_execute.Add(self.m_clusterParam, 0, wx.ALL, 10)        
        self.m_refresh = wx.Button(panel, -1, "Execute")
        self.m_refresh.Bind(wx.EVT_BUTTON, self.OnExecute)
        hor_box_execute.Add(self.m_refresh, 0, wx.ALL, 10)
        vert_box.Add(hor_box_execute, 0, wx.ALL, 2)

        vert_box_MP = wx.BoxSizer(wx.VERTICAL)
        vert_box_MP_Name = wx.BoxSizer(wx.VERTICAL)
        vert_box_Next_P = wx.BoxSizer(wx.VERTICAL)
        m_textM1 = wx.StaticText(panel, -1, "Clusters:")
        vert_box_MP.Add(m_textM1, 0, wx.ALL, 4)
        self.m_addMP = wx.Button(panel, -1, "Rename")
        self.m_addMP.Bind(wx.EVT_BUTTON, self.OnRenameModelPage)
        vert_box_MP.Add(self.m_addMP, 0, wx.ALL, 4)
        self.m_delMP = wx.Button(panel, -1, "Delete")
        self.m_delMP.Bind(wx.EVT_BUTTON, self.OnDelModelPage)
        vert_box_MP.Add(self.m_delMP, 0, wx.ALL, 4)
        hor_box_modelPs.Add(vert_box_MP, 0, wx.ALL, 10)
        
        self.m_modelPages = []
        self.m_modelPagesList = wx.ListBox(panel, size=(150,110)) # choices = mp_names, 
        self.m_modelPagesList.Bind(wx.EVT_LISTBOX, self.OnSelectModelPage)
        hor_box_modelPs.Add(self.m_modelPagesList, 0, wx.ALL, 10)
        
        m_textM2 = wx.StaticText(panel, -1, "Selected\nModel\nPage of\nCluster:")
        hor_box_modelPs.Add(m_textM2, 0, wx.ALL, 16)
        m_textM3 = wx.StaticText(panel, -1, "Cluster / Model Page Name")
        vert_box_MP_Name.Add(m_textM3, 0, wx.ALL, 6)
        self.m_modelPageName = wx.TextCtrl(panel, -1, "(name)", size=(170,-1))
        vert_box_MP_Name.Add(self.m_modelPageName, 0, wx.ALL, 2)        
        self.m_modelPageDoc = wx.TextCtrl(panel, -1, "(doc, pg)", size=(170,-1), style=wx.TE_READONLY)
        vert_box_MP_Name.Add(self.m_modelPageDoc, 0, wx.ALL, 2)        
        hor_box_modelPs.Add(vert_box_MP_Name, 0, wx.ALL, 10)

        m_textM6 = wx.StaticText(panel, -1, "For each cluster the central Data Page will be a Model Page for that cluster.\nAfter you have defined the desired Model Pages, "+
                                 "the program tries\nto find more of similar pages in the next phase.")
        vert_box_Next_P.Add(m_textM6, 0, wx.ALL, 10)
        self.m_nextPhase = wx.Button(panel, -1, "Next >>")
        self.m_nextPhase.Bind(wx.EVT_BUTTON, self.OnNextPhase)
        vert_box_Next_P.Add(self.m_nextPhase, 0, wx.ALL, 10)
        hor_box_modelPs.Add(vert_box_Next_P, 0, wx.ALL, 10)
        vert_box.Add(hor_box_modelPs, 0, wx.ALL, 6)

        m_textE1 = wx.StaticText(panel, -1, "You may wish to change the Model Page of the selected cluster.\nSelect a Data Page from the drop-down below, and check 'Make This Page...'. ")
        hor_box_isModel.Add(m_textE1, 0, wx.ALL, 10)        
        self.m_isModelPage = wx.CheckBox(panel, label="Make This Page\nthe Model Page")        
        self.Bind(wx.EVT_CHECKBOX, self.OnIsModelPage) 
        hor_box_isModel.Add(self.m_isModelPage, 0, wx.ALL, 6)
        m_textE2 = wx.StaticText(panel, -1, "You may wish to eliminate this whole cluster\nfrom the results: Check 'Exclude This Cluster'. ")
        hor_box_isModel.Add(m_textE2, 0, wx.ALL, 10)        
        self.m_isExcluded = wx.CheckBox(panel, label="Exclude\nThis Cluster")        
        self.Bind(wx.EVT_CHECKBOX, self.OnIsExcluded) 
        hor_box_isModel.Add(self.m_isExcluded, 0, wx.ALL, 6)
        vert_box.Add(hor_box_isModel, 0, wx.ALL, 6)

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


    def OnClusterAlgo(self, event):
        algo = self.m_clustAlgo.GetString(self.m_clustAlgo.GetSelection());
        if algo=='K-Means':
            self.m_textClustParam.SetLabel('Number of\nClusters')
            self.m_clusterParam.SetValue('12')
        elif algo=='DBSCAN':
            self.m_textClustParam.SetLabel('Epsilon')
            self.m_clusterParam.SetValue('0.5')
        elif algo=='AffinityPropagation':
            self.m_textClustParam.SetLabel('Damping')
            self.m_clusterParam.SetValue('0.5')
        elif algo=='SpectralClustering':
            self.m_textClustParam.SetLabel('Number of\nClusters')
            self.m_clusterParam.SetValue('12')
        else:
            resp = wx.MessageBox('Precomp first', "You must first press the 'Precomp' button", wx.OK | wx.ICON_WARNING)
    

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
        self.m_extracted.SetLabel(extrText)
        if self.m_docPage:
            self.OnShowPng(event)


    # Return first sample of the given cluster
    def GetFirstDocIndexAndPageNumber(self, cluster):
        for i,p in enumerate(self.m_labels):
            if p==cluster:
                line = self.m_links1[i]
                pg = int(line.split("#page=")[1])
                file1 = line.rfind(os.path.sep)
                file2 = line.rfind('#')
                file = line[file1+1:file2]
                for j,f in enumerate(files_to_process):
                    if file in f:
                        return j, pg 
        return -1, -1

    
    # Return the sample of the given cluster that has the minimum distance sum (to other samples of that cluster) = central sample
    def GetBestDocIndexAndPageNumber(self, cluster):
        min_sum = -1.0
        min_idx = -1
        for i,pi in enumerate(self.m_labels):
            if pi==cluster:
                dist_sum = 0.0
                for j,pj in enumerate(self.m_labels):
                    if j==i:
                        continue
                    if pj==cluster:
                        dist_sum += self.m_dist_matrix[i][j]
                if min_sum<0.0:
                    min_sum = dist_sum
                    min_idx = i
                elif min_sum<dist_sum:
                    min_sum = dist_sum
                    min_idx = i

        line = self.m_links1[min_idx]
        pg = int(line.split("#page=")[1])
        file1 = line.rfind(os.path.sep)
        file2 = line.rfind('#')
        file = line[file1+1:file2]
        for k,f in enumerate(files_to_process):
            if file in f:
                return k, pg 
                
        return -1, -1
    

    def GetDocIndexAndPageNumberByLine(self, line):
        pg = int(line.split("#page=")[1])
        file1 = line.rfind(os.path.sep)
        file2 = line.rfind('#')
        file = line[file1+1:file2]
        for j,f in enumerate(files_to_process):
            if file in f:
                return j, pg 
        return -1, -1
    

    def OnPrecompute(self, event):
        if self.m_dist_matrix is not None:
            resp = wx.MessageDialog(None,
                                    "Distance Matrix has been computed already.\nAre you sure you want to recompute it?", 
                                    'Distance Matrix Exists', 
                                    wx.YES_NO | wx.ICON_WARNING | wx.NO_DEFAULT
                                    ).ShowModal()
            if resp != wx.ID_YES:
                return

        self.m_dataset_data = []
        prevFile = ''
        self.m_extracted.Clear()
        self.m_extracted.AppendText('Loading page data...\n')
        self.m_extracted.Refresh()
        self.m_extracted.Update()
        for docNumber, docPath in enumerate(self.m_links1):
        
            linkLine = docPath.split('#')
            if len(linkLine)<2:
                continue
            filename = linkLine[0][8:] # remove 'file:///'
            if filename!=prevFile:
                doc = fitz.open(filename)
                prevFile = filename
        
            for pageNumber, page in enumerate(doc):
                page_n = int(linkLine[1][5:])
                if page_n==pageNumber+1:
                    # extracting text from page
                    pageText = page.get_text('text')
                    self.m_dataset_data.append(pageText)
                    self.m_extracted.AppendText('file: '+filename+', page: '+str(page_n)+'\n')
                            
        self.m_extracted.AppendText('Loaded '+str(len(self.m_dataset_data))+' pages of data\n')

        self.m_extracted.AppendText('Computing distance matrix...\n')
        self.m_extracted.Refresh()
        self.m_extracted.Update()
        dim = len(self.m_dataset_data)
        self.m_dist_matrix = np.zeros((dim, dim))
        for i, col in enumerate(self.m_dataset_data):
            if i>0 and (i % (((i-1)//100+1)*10)) == 0:
                self.m_extracted.AppendText(' ' + str(i) + '/' + str(dim) + '\n')
                self.m_extracted.Refresh()
                self.m_extracted.Update()
            for j, row in enumerate(self.m_dataset_data):
                if j<i:
                    pass
                elif i==j:
                    self.m_dist_matrix[i][j] = 0.0
                else:
                    words_i = self.m_dataset_data[i].split('\n')
                    nbrWords_i = len(words_i)
                    counts_i = Counter(words_i)
                    words_j = self.m_dataset_data[j].split('\n')
                    nbrWords_j = len(words_j)
                    counts_j = Counter(words_j)
                    diff = computeDifference(nbrWords_i, counts_i, nbrWords_j, counts_j)
                    self.m_dist_matrix[i][j] = diff
                    self.m_dist_matrix[j][i] = diff
                    
        self.m_extracted.AppendText('Data loaded and distance matrix computed!\n')
        self.m_extracted.Refresh()
        self.m_extracted.Update()
        
        self.m_clustAlgo.Clear()
        self.m_clustAlgo.AppendItems(['K-Means', 'DBSCAN', 'AffinityPropagation', 'SpectralClustering'])
    
    
    def OnExecute(self, event):
        self.m_extracted.Clear()
        algo = self.m_clustAlgo.GetString(self.m_clustAlgo.GetSelection());
        self.m_labels = None
        
        if algo=='K-Means':
            
            nbr_str = self.m_clusterParam.GetValue()
            try:
                nbr_clusters = int(nbr_str)
                if nbr_clusters<=0:
                    resp = wx.MessageBox("Parameter NbrClusters must be positive", 'No valid NbrClusters entered', wx.OK | wx.ICON_WARNING)
                    return
            except ValueError:
                resp = wx.MessageBox("Parameter NbrClusters is not a integer number", 'No valid NbrClusters entered', wx.OK | wx.ICON_WARNING)
                return
            
            vectorizer = TfidfVectorizer(
                max_df=1.0,
                min_df=1,
                stop_words="english",
            )
            X_tfidf = vectorizer.fit_transform(self.m_dataset_data)            
            self.m_extracted.AppendText("n_samples: "+str(X_tfidf.shape[0])+", n_features: "+str(X_tfidf.shape[1])+"\n")
            
            kmeans = KMeans(
                n_clusters=nbr_clusters,
                max_iter=100,
                n_init=5
            ).fit(X_tfidf)
            cluster_ids, cluster_sizes = np.unique(kmeans.labels_, return_counts=True)
            self.m_extracted.AppendText("Number of elements asigned to each cluster: "+str(cluster_sizes)+"\n")
            self.m_labels = kmeans.labels_
            
            # self.m_extracted.AppendText('Labels: ' + str(self.m_labels) + '\n')
            # self.m_extracted.AppendText('Pages in each cluster: ' + str(Counter(self.m_labels)) + '\n')
            
            str_line = ""
            for j, c in enumerate(cluster_sizes):
                str_line = "Cluster " + str(j) + ", pages = "
                for i, p in enumerate(self.m_labels):
                    if p==j:
                        str_line += str(i) + " "
                self.m_extracted.AppendText(str_line + '\n')
                
            mp_names = []
            for j, c in enumerate(cluster_sizes):
                mp_names.append("Cluster " + str(j) + " (" + str(np.count_nonzero(self.m_labels==j)) + " pages)")
            self.m_modelPagesList.Clear()
            self.m_modelPagesList.AppendItems(mp_names)
            
            self.m_modelPages.clear()
            self.m_modelPages = [{ "name": mp, "doc_index": -1, "page_nbr": -1, "excluded": False } for mp in mp_names]
            for i, mp in enumerate(self.m_modelPages):
                # TODO: First => Best
                doc, pg = self.GetBestDocIndexAndPageNumber(i)
                mp["doc_index"] = doc
                mp["page_nbr"] = pg
                    
            self.m_choice.Clear()
            self.m_choice.AppendItems(self.m_links2)
            if self.m_choice.GetCount()>0:
                self.m_choice.SetSelection(0)
                self.OnChoice(None)

            '''            
            # K-Means with LCA starts here...
            lsa = make_pipeline(TruncatedSVD(n_components=100), Normalizer(copy=False))
            X_lsa = lsa.fit_transform(X_tfidf)
            explained_variance = lsa[0].explained_variance_ratio_.sum()
            
            self.m_extracted.AppendText("Explained variance of the SVD step: "+str(explained_variance * 100.0)+"\n")
            kmeans_lca = KMeans(
                n_clusters=nbr_clusters,
                max_iter=100,
                n_init=1
            ).fit(X_lsa)
            
            cluster_ids, cluster_sizes = np.unique(kmeans_lca.labels_, return_counts=True)
            self.m_extracted.AppendText("LCA: Number of elements asigned to each cluster: "+str(cluster_sizes)+"\n")
            self.m_labels = kmeans_lca.labels_
            
            self.m_extracted.AppendText('Labels: ' + str(self.m_labels) + '\n')
            self.m_extracted.AppendText('Pages in each cluster: ' + str(Counter(self.m_labels)) + '\n')
            
            original_space_centroids = lsa[0].inverse_transform(kmeans_lca.cluster_centers_)
            order_centroids = original_space_centroids.argsort()[:, ::-1]
            terms = vectorizer.get_feature_names() # Should be: vectorizer.get_feature_names_out()
            
            for i in range(nbr_clusters):
                clus_str = 'Cluster ' + str(i) + ':'
                for ind in order_centroids[i, :10]:
                    clus_str += terms[ind] + ' '
                self.m_extracted.AppendText(clus_str + '\n')
            '''            
            
        elif algo=='DBSCAN':
            
            eps_str = self.m_clusterParam.GetValue()
            try:
                epsilon = float(eps_str)
                if epsilon<=0.0:
                    resp = wx.MessageBox("Parameter Epsilon must be positive", 'No valid Epsilon entered', wx.OK | wx.ICON_WARNING)
                    return
            except ValueError:
                resp = wx.MessageBox("Parameter Epsilon is not a floating-point number", 'No valid Epsilon entered', wx.OK | wx.ICON_WARNING)
                return
                
            self.m_extracted.Clear()
            db = DBSCAN(eps=epsilon, metric='precomputed', min_samples=3).fit(self.m_dist_matrix)
            self.m_labels = db.labels_
            
            # self.m_extracted.AppendText('Labels: ' + str(self.m_labels) + '\n')
            self.m_extracted.AppendText('Pages in each cluster: ' + str(Counter(self.m_labels)) + '\n')
            # Then take a look at the classification results!
            n_clusters_ = len(set(self.m_labels)) - (1 if -1 in self.m_labels else 0) # Number of classes
            self.m_extracted.AppendText('The number of clusters is: ' + str(n_clusters_) + '\n')
            self.m_extracted.Refresh()
            self.m_extracted.Update()
            
            str_line = ""
            for j in range(n_clusters_):
                str_line = "Cluster " + str(j) + ", pages = "
                for i, p in enumerate(self.m_labels):
                    if p==j:
                        str_line += str(i) + " "
                self.m_extracted.AppendText(str_line + '\n')
            self.m_extracted.AppendText(str(np.count_nonzero(self.m_labels==-1)) + ' pages outside clusters\n')
            
            mp_names = []
            for j in range(n_clusters_):
                mp_names.append("Cluster " + str(j) + " (" + str(np.count_nonzero(self.m_labels==j)) + " pages)")
                
            self.m_modelPages.clear()
            self.m_modelPages = [{ "name": mp, "doc_index": -1, "page_nbr": -1, "excluded": False } for mp in mp_names]
            for i, mp in enumerate(self.m_modelPages):
                # TODO: First => Best
                doc, pg = self.GetBestDocIndexAndPageNumber(i)
                mp["doc_index"] = doc
                mp["page_nbr"] = pg
                
            mp_names.append("Left-overs (" + str(np.count_nonzero(self.m_labels==-1)) + " pages)")
            self.m_modelPagesList.Clear()
            self.m_modelPagesList.AppendItems(mp_names)

        elif algo=='AffinityPropagation':

            damp_str = self.m_clusterParam.GetValue()
            try:
                damp = float(damp_str)
                if damp<0.5 or damp>=1.0:
                    resp = wx.MessageBox("Parameter Damping must be in range [0.5, 1.0)", 'No valid Damping entered', wx.OK | wx.ICON_WARNING)
                    return
            except ValueError:
                resp = wx.MessageBox("Parameter Damping is not a floating-point number", 'No valid Damping entered', wx.OK | wx.ICON_WARNING)
                return
                
            self.m_extracted.Clear()
            # https://stackoverflow.com/questions/35494458/affinity-propagation-in-python
            aff_matrix = 1 - self.m_dist_matrix
            ap = AffinityPropagation(damping=damp, affinity='precomputed', random_state=None).fit(aff_matrix)
            self.m_labels = ap.labels_
            
            # self.m_extracted.AppendText('Labels: ' + str(self.m_labels) + '\n')
            self.m_extracted.AppendText('Pages in each cluster: ' + str(Counter(self.m_labels)) + '\n')
            # Then take a look at the classification results!
            n_clusters_ = len(set(self.m_labels)) - (1 if -1 in self.m_labels else 0) # Number of classes
            self.m_extracted.AppendText('The number of clusters is: ' + str(n_clusters_) + '\n')
            self.m_extracted.Refresh()
            self.m_extracted.Update()
            
            str_line = ""
            for j in range(n_clusters_):
                str_line = "Cluster " + str(j) + ", pages = "
                for i, p in enumerate(self.m_labels):
                    if p==j:
                        str_line += str(i) + " "
                self.m_extracted.AppendText(str_line + '\n')
            self.m_extracted.AppendText(str(np.count_nonzero(self.m_labels==-1)) + ' pages outside clusters\n')
            
            mp_names = []
            for j in range(n_clusters_):
                mp_names.append("Cluster " + str(j) + " (" + str(np.count_nonzero(self.m_labels==j)) + " pages)")
                
            self.m_modelPages.clear()
            self.m_modelPages = [{ "name": mp, "doc_index": -1, "page_nbr": -1, "excluded": False } for mp in mp_names]
            for i, mp in enumerate(self.m_modelPages):
                # TODO: First => Best
                doc, pg = self.GetBestDocIndexAndPageNumber(i)
                mp["doc_index"] = doc
                mp["page_nbr"] = pg
                
            mp_names.append("Left-overs (" + str(np.count_nonzero(self.m_labels==-1)) + " pages)")
            self.m_modelPagesList.Clear()
            self.m_modelPagesList.AppendItems(mp_names)
            
        elif algo=='SpectralClustering':

            nbr_str = self.m_clusterParam.GetValue()
            try:
                nbr_clusters = int(nbr_str)
                if nbr_clusters<=0:
                    resp = wx.MessageBox("Parameter NbrClusters must be positive", 'No valid NbrClusters entered', wx.OK | wx.ICON_WARNING)
                    return
            except ValueError:
                resp = wx.MessageBox("Parameter NbrClusters is not a integer number", 'No valid NbrClusters entered', wx.OK | wx.ICON_WARNING)
                return
                
            self.m_extracted.Clear()
            # https://stackoverflow.com/questions/33773916/precomputed-distances-for-spectral-clustering-with-scikit-learn
            delta = self.m_dist_matrix.max()-self.m_dist_matrix.min()
            sim_matrix = np.exp(- self.m_dist_matrix ** 2 / (2.0 * delta ** 2))
            sc = SpectralClustering(n_clusters=nbr_clusters, affinity='precomputed').fit(sim_matrix)
            self.m_labels = sc.labels_
            
            # self.m_extracted.AppendText('Labels: ' + str(self.m_labels) + '\n')
            self.m_extracted.AppendText('Pages in each cluster: ' + str(Counter(self.m_labels)) + '\n')
            # Then take a look at the classification results!
            n_clusters_ = len(set(self.m_labels)) - (1 if -1 in self.m_labels else 0) # Number of classes
            self.m_extracted.AppendText('The number of clusters is: ' + str(n_clusters_) + '\n')
            self.m_extracted.Refresh()
            self.m_extracted.Update()
            
            str_line = ""
            for j in range(n_clusters_):
                str_line = "Cluster " + str(j) + ", pages = "
                for i, p in enumerate(self.m_labels):
                    if p==j:
                        str_line += str(i) + " "
                self.m_extracted.AppendText(str_line + '\n')
            self.m_extracted.AppendText(str(np.count_nonzero(self.m_labels==-1)) + ' pages outside clusters\n')
            
            mp_names = []
            for j in range(n_clusters_):
                mp_names.append("Cluster " + str(j) + " (" + str(np.count_nonzero(self.m_labels==j)) + " pages)")
                
            self.m_modelPages.clear()
            self.m_modelPages = [{ "name": mp, "doc_index": -1, "page_nbr": -1, "excluded": False } for mp in mp_names]
            for i, mp in enumerate(self.m_modelPages):
                # TODO: First => Best
                doc, pg = self.GetBestDocIndexAndPageNumber(i)
                mp["doc_index"] = doc
                mp["page_nbr"] = pg
                
            mp_names.append("Left-overs (" + str(np.count_nonzero(self.m_labels==-1)) + " pages)")
            self.m_modelPagesList.Clear()
            self.m_modelPagesList.AppendItems(mp_names)
            
        else:
            resp = wx.MessageBox("You must first select a clustering algorithm\nBefore that, press the 'Precomp' button", 'Select algorithm first', wx.OK | wx.ICON_WARNING)

        if self.m_labels is not None:
            allLogText = []
            nbrClusters = self.m_modelPagesList.GetCount()
            for cl in range(nbrClusters):
                if cl<len(self.m_modelPages):
                    doc_index = self.m_modelPages[cl]["doc_index"]
                    page_nbr = self.m_modelPages[cl]["page_nbr"]
                    str_doc_pg = "Doc " + str(doc_index) + ", Pg " + str(page_nbr)
                    allLogText.append(self.m_modelPagesList.GetString(cl) + ': ' + str_doc_pg + ' = ' + files_to_process[doc_index] + '\n')
                else:
                    allLogText.append(self.m_modelPagesList.GetString(cl) + '\n')                    
                for i,p in enumerate(self.m_labels):
                    if p==cl:
                        allLogText.append(self.m_links1[i])
                allLogText.append("\n")
            dlg = wx.MessageDialog(self, 
                "Clusters data are available in memory.\nDo you want to Copy the data to Clipboard?",
                "Copy Clusters?", wx.YES_NO|wx.NO_DEFAULT|wx.ICON_QUESTION)
            result = dlg.ShowModal()
            dlg.Destroy()
            if result == wx.ID_YES:
                dataObj = wx.TextDataObject()
                dataObj.SetText("".join(allLogText))
                wx.TheClipboard.Open()
                wx.TheClipboard.SetData(dataObj)
                wx.TheClipboard.Close()
        

    def LoadResults(self):
        '''
        input_file = resultsFolder2 + "/Phase2_results.txt"
        if not os.path.exists(input_file):
            return
        with open(input_file, "r") as file:
            self.m_data2 = file.readlines()

        self.m_links2 = []
        for d in self.m_data2:
            if "file:///" in d:
                self.m_links2.append(d)
                
        self.m_choice.Clear()
        self.m_choice.AppendItems(self.m_links2)
        if self.m_choice.GetCount()>0:
            self.m_choice.SetSelection(0)
            self.OnChoice(None)
        '''
        self.m_links2 = []
        if self.m_prevChoice.GetCount()>0:
            self.m_prevChoice.SetSelection(0)
            self.OnPrevChoice(None)
        pass

    
    def OnNextPhase(self, event):
        output_file = resultsFolder3 + "/Phase3_modelPages.json"
        jsonString = json.dumps(self.m_modelPages)
        jsonFile = open(output_file, "w")
        jsonFile.write(jsonString)
        jsonFile.close()        

        os.system('python VesselAI_extractData_3.py')


    def OnSelectModelPage(self, event):
        idx = self.m_modelPagesList.GetSelection()
        if idx>=0:
            mpage = self.m_modelPagesList.GetString(idx);
            self.m_modelPageName.SetValue(mpage)
            self.m_links2.clear()
            for i,p in enumerate(self.m_labels):
                if p==idx:
                    self.m_links2.append(self.m_links1[i])
                    
            self.m_choice.Clear()
            self.m_choice.AppendItems(self.m_links2)

            if idx<len(self.m_modelPages):
                doc_index = self.m_modelPages[idx]["doc_index"]
                page_nbr = self.m_modelPages[idx]["page_nbr"]
                str_doc_pg = "Doc " + str(doc_index) + ", Pg " + str(page_nbr)
                self.m_modelPageDoc.SetValue(str_doc_pg)
                if os.path.sep=='/':
                    file_part = files_to_process[doc_index]
                else:
                    file_part = files_to_process[doc_index].replace('/', os.path.sep)
                page_part = "#page="+str(page_nbr)
                for i,p in enumerate(self.m_links2):
                    if file_part in p and page_part in p:
                        self.m_choice.SetSelection(i)
                        self.OnChoice(None)
                        self.m_isModelPage.SetValue(True)
                        break
                else:
                    if self.m_choice.GetCount()>0:
                        self.m_choice.SetSelection(0)
                        self.OnChoice(None)
                    doc, pg = self.GetFirstDocIndexAndPageNumber(idx)
                    if doc==self.m_modelPages[idx]["doc_index"] and pg==self.m_modelPages[idx]["page_nbr"]:
                        self.m_isModelPage.SetValue(True)
                    else:
                        self.m_isModelPage.SetValue(False)
                self.m_isExcluded.SetValue(self.m_modelPages[idx]["excluded"])
            else:
                self.m_modelPageDoc.SetValue("")
                self.m_isModelPage.SetValue(False)
                self.m_isExcluded.SetValue(False)
        else:
            self.m_modelPageName.SetValue("")
            self.m_modelPageDoc.SetValue("")


    def OnRenameModelPage(self, event):
        name = self.m_modelPageName.GetValue()
        idx = self.m_modelPagesList.GetSelection()
        if not name.strip():
            resp = wx.MessageBox('No Name', 'Operation cancelled', wx.OK | wx.ICON_WARNING)
            return            
        if idx<0 or idx>=len(self.m_modelPages):
            resp = wx.MessageBox('No Model Page selected', 'Operation cancelled', wx.OK | wx.ICON_WARNING)
            return            
        
        self.m_modelPages[idx]["name"] = name.strip()
            
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

    
    def OnIsModelPage(self, event):
        idx = self.m_modelPagesList.GetSelection()
        if idx<0 or idx>=len(self.m_modelPages):
            return
        check = self.m_isModelPage.IsChecked()
        if check:
            line = self.m_choice.GetString(self.m_choice.GetSelection());
            doc, pg = self.GetDocIndexAndPageNumberByLine(line)
            self.m_modelPages[idx]["doc_index"] = doc 
            self.m_modelPages[idx]["page_nbr"] = pg
            str_doc_pg = "Doc " + str(doc) + ", Pg " + str(pg)
            self.m_modelPageDoc.SetValue(str_doc_pg)
        else:
            resp = wx.MessageBox("You cannot clear this setting\nInstead, set this for another page\n(selected thru lower drop-down)", 'Cannot Clear', wx.OK | wx.ICON_WARNING)
            self.m_isModelPage.SetValue(True)


    def OnIsExcluded(self, event):
        idx = self.m_modelPagesList.GetSelection()
        if idx<0 or idx>=len(self.m_modelPages):
            return
        check = self.m_isExcluded.IsChecked()
        self.m_modelPages[idx]["excluded"] = check
    

    def OnOpenLink(self, event):
        url = self.m_choice.GetString(self.m_choice.GetSelection())
        webbrowser.open(url)

    def OnCopyLink(self, event):
        url = self.m_choice.GetString(self.m_choice.GetSelection())
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(url))
            wx.TheClipboard.Close()

    def OnShowPng(self, event, folder1=True):
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
            for d in self.m_data1:
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
            extrText = extrText[0:findOpeningBrace] + extrText[findClosingBrace+1:]
        self.m_extracted.SetLabel(extrText)
            
        idx = self.m_modelPagesList.GetSelection()
        if idx>=0 and idx<len(self.m_modelPages):
            doc, pg = self.GetDocIndexAndPageNumberByLine(line)
            if doc==self.m_modelPages[idx]["doc_index"] and pg==self.m_modelPages[idx]["page_nbr"]:
                self.m_isModelPage.SetValue(True)
            else:
                self.m_isModelPage.SetValue(False)
            self.m_isExcluded.SetValue(self.m_modelPages[idx]["excluded"])
        else:
            self.m_isModelPage.SetValue(False)
            self.m_isExcluded.SetValue(False)
            
        if self.m_docPage:
            self.OnShowPng(event)


app = wx.App(redirect=True)   # Error messages go to popup window
top = Frame("VesselAI_ExtractData_Phase_2_Clusters")
top.Show()
app.MainLoop()