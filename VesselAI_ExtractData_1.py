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


aboutText = """<p>Sorry, there is no information about this program. It is
running on version %(wxpy)s of <b>wxPython</b> and %(python)s of <b>Python</b>.
See <a href="http://wiki.wxpython.org">wxPython Wiki</a></p>""" 

resultsFolder1 = "Phase_1_Docs_Pages"

# We have 8 groups of keywords and other words, see about line 330
        # Fuel consump: key = g/kWh OR kJ/kWh, other = fuel and consumption
        # Fuel SFOC: key = g/kWh OR kJ/kWh, other = sfoc
        # Fuel BSFC: key = g/kWh OR kJ/kWh, other = bsfc
        # Charge air: key = kg/s OR kg/kWh OR m3/h, other = charge air flow
        # Combust air: key = kg/s OR kg/kWh OR m3/h, other = combustion air flow
        # Exhaust gas: key = kg/s OR kg/kWh OR m3/h, other = exhaust gas flow
        # Heat balance: key = kW, other = heat balance
        # Heat quantity: key = kW, other = heat quantity

key_words_to_find = [
    "g/kWh kJ/kWh",
    "g/kWh kJ/kWh",    
    "g/kWh kJ/kWh",    
    "kg/s kg/kWh",
    "kg/s kg/kWh",
    "kg/s kg/kWh",
    "kW"
    ]
other_words_to_find = [
    "fuel consumption",
    "sfoc",
    "bsfc",    
    "charge air flow",
    "combustion air flow",    
    "exhaust gas flow",    
    "heat balance"
    ]

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


# The following 'global' functions have been copied from file 'ExtractInfoPages_AllLines.py'.
# That program just processed the input files ('files_to_process') and extracted the relevant pages.
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
        wx.Dialog.__init__(self, None, -1, "About VesselAI_ExtractData_Phase_1_KeyWords",
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
        hor_box_search = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_dataSets = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_refresh = wx.BoxSizer(wx.HORIZONTAL)
        hor_box_exclude = wx.BoxSizer(wx.HORIZONTAL)
        
        self.m_docPage = ""
        # file:///<file_path>.pdf#page=<page_number> entries will be added to this array for those pages where 'Exclude This Page from Results' check box is set.
        self.m_excludedPages = [] 
        input_file = resultsFolder1 + "/Phase1_excluded.txt"
        if os.path.exists(input_file):
            with open(input_file, 'r') as fp:
                self.m_excludedPages = fp.readlines()

        m_text0 = wx.StaticText(panel, -1, "  Enter Keywords and Other words. Add/Mod: Adds a new item to the end of the list, if Name is new; otherwise, it Modifies existing item.\n"+
                                "  Press 'Refresh' to see what data pages match these criteria => they appear in the drop-down below. Select the desired data page from the drop-down to see its contents.")
        vert_box.Add(m_text0, 0, wx.ALL, 10)
        
        vert_box_DS = wx.BoxSizer(wx.VERTICAL)
        vert_box_DS_Name = wx.BoxSizer(wx.VERTICAL)
        vert_box_DS_Keyw = wx.BoxSizer(wx.VERTICAL)
        vert_box_DS_Othr = wx.BoxSizer(wx.VERTICAL)
        m_textD1 = wx.StaticText(panel, -1, "Data Sets:")
        vert_box_DS.Add(m_textD1, 0, wx.ALL, 4)
        self.m_addDS = wx.Button(panel, -1, "Add/Mod")
        self.m_addDS.Bind(wx.EVT_BUTTON, self.OnAddDataSet)
        vert_box_DS.Add(self.m_addDS, 0, wx.ALL, 4)
        self.m_delDS = wx.Button(panel, -1, "Delete")
        self.m_delDS.Bind(wx.EVT_BUTTON, self.OnDelDataSet)
        vert_box_DS.Add(self.m_delDS, 0, wx.ALL, 4)
        hor_box_dataSets.Add(vert_box_DS, 0, wx.ALL, 10)
        
        # Fuel consump: key = g/kWh OR kJ/kWh, other = fuel consumption
        # Fuel SFOC: key = g/kWh OR kJ/kWh, other = sfoc
        # Fuel BSFC: key = g/kWh OR kJ/kWh, other = bsfc
        # Charge air: key = kg/s OR kg/kWh OR m3/h, other = charge air flow
        # Combust air: key = kg/s OR kg/kWh OR m3/h, other = combustion air flow
        # Exhaust gas: key = kg/s OR kg/kWh OR m3/h, other = exhaust gas flow
        # Heat balance: key = kW, other = heat balance
        # Heat quantity: key = kW, other = heat quantity
        # Heat rad engine: key = kW, other = heat radiation engine
        # Charge air cool: key = kW, other = charge air cooler
        # Lube oil cool: key = kW, other = lube oil cooler
        # Jacket cool: key = kW, other = jacket cooling
        self.m_dataSets = [
            { "name": "Fuel consump", "keywords": "g/kWh kJ/kWh", "other_words": "fuel consumption"},
            { "name": "Fuel SFOC", "keywords": "g/kWh kJ/kWh", "other_words": "sfoc"},
            { "name": "Fuel BSFC", "keywords": "g/kWh kJ/kWh", "other_words": "bsfc"},
            { "name": "Charge air", "keywords": "kg/s kg/kWh m3/h", "other_words": "charge air flow"},
            { "name": "Combust air", "keywords": "kg/s kg/kWh m3/h", "other_words": "combustion air flow"},
            { "name": "Exhaust gas", "keywords": "kg/s kg/kWh m3/h", "other_words": "exhaust gas flow"},
            { "name": "Heat balance", "keywords": "kW", "other_words": "heat balance"},
            { "name": "Heat quantity", "keywords": "kW", "other_words": "heat quantity"},
            { "name": "Heat rad engine", "keywords": "kW", "other_words": "heat radiation engine"},
            { "name": "Charge air cool", "keywords": "kW", "other_words": "charge air cooler"},
            { "name": "Lube oil cool", "keywords": "kW", "other_words": "lube oil cooler"},
            { "name": "Jacket cool", "keywords": "kW", "other_words": "jacket cooling"},
            ]
        
        input_file_j = resultsFolder1 + "/Phase1_dataSets.json"
        try:
            fileJson = open(input_file_j, "r")
            jsonContent = fileJson.read()
            self.m_dataSets = json.loads(jsonContent)
        except IOError:
            pass # using the default self.m_dataSets data defined above...
            
        ds_names = [i["name"] for i in self.m_dataSets]
        self.m_dataSetsList = wx.ListBox(panel, choices = ds_names, size=(100,100))
        self.m_dataSetsList.Bind(wx.EVT_LISTBOX, self.OnSelectDataSet)
        hor_box_dataSets.Add(self.m_dataSetsList, 0, wx.ALL, 10)
        
        m_textD2 = wx.StaticText(panel, -1, "Selected\nData Set:")
        hor_box_dataSets.Add(m_textD2, 0, wx.ALL, 16)
        m_textD3 = wx.StaticText(panel, -1, "Name")
        vert_box_DS_Name.Add(m_textD3, 0, wx.ALL, 6)
        self.m_dataSetName = wx.TextCtrl(panel, -1, "(name)", size=(100,-1))
        vert_box_DS_Name.Add(self.m_dataSetName, 0, wx.ALL, 2)        
        hor_box_dataSets.Add(vert_box_DS_Name, 0, wx.ALL, 10)        
        m_textD4 = wx.StaticText(panel, -1, "Keywords")
        vert_box_DS_Keyw.Add(m_textD4, 0, wx.ALL, 6)
        self.m_dataSetKeywords = wx.TextCtrl(panel, -1, "(keywords)", size=(150,-1))
        vert_box_DS_Keyw.Add(self.m_dataSetKeywords, 0, wx.ALL, 2)        
        hor_box_dataSets.Add(vert_box_DS_Keyw, 0, wx.ALL, 10)        
        m_textD5 = wx.StaticText(panel, -1, "Other words")
        vert_box_DS_Othr.Add(m_textD5, 0, wx.ALL, 6)
        self.m_dataSetOtherWords = wx.TextCtrl(panel, -1, "(other words)", size=(200,-1))
        vert_box_DS_Othr.Add(self.m_dataSetOtherWords, 0, wx.ALL, 2)        
        hor_box_dataSets.Add(vert_box_DS_Othr, 0, wx.ALL, 10)        
        m_textD6 = wx.StaticText(panel, -1, "Name is an ID for the Data set.\nKeywords are OR'ed, but must be whole table cell.\nOther words are AND'ed, they all must appear.\n"+
                                 "Case is respected for Keywords.\nCase is ignored for Other words.")
        hor_box_dataSets.Add(m_textD6, 0, wx.ALL, 10)
        
        vert_box.Add(hor_box_dataSets, 0, wx.ALL, 4)

        m_textR0 = wx.StaticText(panel, -1, "After you have entered the desired criteria to 'Data Sets' above, press button 'Refresh', " + 
                                "and the drop-down below is filled with matching data pages.\n" +
                                "Check that the desired data pages are indeed in the drop-down (use check box Exclude...), then press button 'Next >>', " +
                                "and the Phase 2 window is opened.")
        vert_box.Add(m_textR0, 0, wx.ALL, 6)
        
        m_textR1 = wx.StaticText(panel, -1, "Good data pages usually contain numbers;\nto enforce this use the parameter.")
        hor_box_refresh.Add(m_textR1, 0, wx.ALL, 4)        
        m_textR2 = wx.StaticText(panel, -1, "Min Nbrs in Page\nto accept page")
        hor_box_refresh.Add(m_textR2, 0, wx.ALL, 4)
        self.m_numbersInPage = wx.TextCtrl(panel, -1, "0", size=(50,-1))
        hor_box_refresh.Add(self.m_numbersInPage, 0, wx.ALL, 4)
        
        self.m_refresh = wx.Button(panel, -1, "Refresh")
        self.m_refresh.Bind(wx.EVT_BUTTON, self.OnRefresh)
        hor_box_refresh.Add(self.m_refresh, 0, wx.ALL, 6)
        self.m_nextPhase = wx.Button(panel, -1, "Next >>")
        self.m_nextPhase.Bind(wx.EVT_BUTTON, self.OnNextPhase)
        hor_box_refresh.Add(self.m_nextPhase, 0, wx.ALL, 6)
        vert_box.Add(hor_box_refresh, 0, wx.ALL, 6)

        m_textE1 = wx.StaticText(panel, -1, "You may wish to eliminate some pages from the results. Select result page from the drop-down below, and check\n'Exclude This Page...'. "+
                                 "If you intend to exclude most of the result pages, button 'Exclude All' checks it for all pages.")
        hor_box_exclude.Add(m_textE1, 0, wx.ALL, 6)        
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
        global key_words_to_find
        global other_words_to_find
        key_words_to_find = [i["keywords"] for i in self.m_dataSets]
        other_words_to_find = [i["other_words"] for i in self.m_dataSets]
        if os.path.isdir(resultsFolder1):
            for i in os.listdir(resultsFolder1):
                os.remove(os.path.join(resultsFolder1, i))
        else:
            os.makedirs(resultsFolder1)
            
        min_nbrs_in_page = float(self.m_numbersInPage.GetValue())
        
        output_file = resultsFolder1 + "/Phase1_results.txt"
        with open(output_file, 'w') as fp:
            
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
                    pageText = page.get_text('text')
                    lines = pageText.split('\n')
                    foundKeys = [False]*len(key_words_to_find)
                    foundOthers = [False]*len(other_words_to_find)
                    linesWithOth = []
                    linesWithKey = []
                    # Other words are AND'ed (they must all appear), keywords are OR'ed (one appearing is enough)
                    for ioth,oth in enumerate(other_words_to_find):
                        oths = oth.split(' ')
                        foundAll = True
                        for o in oths:
                            foundOne = False
                            # OPTION: We accept this page, if other words are found in the previous page; but keywords must be found in this page!
                            for ln in lines: # +prev_lines:
                                if o in ln.lower():
                                    foundOne = True
                            if not foundOne:
                                foundAll = False
                        if foundAll:        
                            foundOthers[ioth] = True
                    for ln in lines:
                        for ikey,key in enumerate(key_words_to_find):
                            keys = key.split(' ')
                            for k in keys:
                                if k==ln or ('('+k+')') in ln or ('['+k+']') in ln or ('in '+k) in ln :
                                    foundKeys[ikey] = True

                    nbrOfNbrs = 0
                    if min_nbrs_in_page>0:
                        for ln in lines:
                            if isNumber(ln):
                                nbrOfNbrs += 1
                        
                    # Keyword and other word must match...
                    match = False
                    for ikey,key in enumerate(foundKeys):
                        if foundKeys[ikey] and foundOthers[ikey]:
                            match = True
                    if match and nbrOfNbrs>=min_nbrs_in_page:
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
                            
                        self.m_extracted.AppendText("   Doc:"+str(docNumber)+", Page:"+str(pageNumber+1)+"\n")
                        print("Doc:"+str(docNumber)+", Page:"+str(pageNumber+1)+" => ", file=fp)
                        
                        for ltp in linesToPrint:
                            try:
                                print(ltp.rstrip(), file=fp)
                            except:
                                pass
            
                    prev_lines = lines[:]
      

    def OnSelectDataSet(self, event):
        idx = self.m_dataSetsList.GetSelection()
        if idx>=0:
            self.m_dataSetName.SetValue(self.m_dataSets[idx]["name"])
            self.m_dataSetKeywords.SetValue(self.m_dataSets[idx]["keywords"])
            self.m_dataSetOtherWords.SetValue(self.m_dataSets[idx]["other_words"])
        else:
            self.m_dataSetName.SetValue("")
            self.m_dataSetKeywords.SetValue("")
            self.m_dataSetOtherWords.SetValue("")
   
    def OnAddDataSet(self, event):
        name = self.m_dataSetName.GetValue()
        keyw = self.m_dataSetKeywords.GetValue()
        othw = self.m_dataSetOtherWords.GetValue()
        if not name.strip():
            resp = wx.MessageBox('No Name', 'Operation cancelled', wx.OK | wx.ICON_WARNING)
            return            
        if not keyw.strip():
            resp = wx.MessageBox('No keyword(s)', 'Operation cancelled', wx.OK | wx.ICON_WARNING)
            return            
        idx = -1
        for ids, ds in enumerate(self.m_dataSets):
            if ds["name"]==name:
                idx = ids
                break
        if idx>=0:
            self.m_dataSets[idx]["keywords"] = keyw.strip()
            self.m_dataSets[idx]["other_words"] = othw.strip().lower()
        else:
            newitem = { "name": name, "keywords": keyw.strip(), "other_words": othw.strip().lower() }
            self.m_dataSets.append(newitem)
            
        ds_names = [i["name"] for i in self.m_dataSets]
        self.m_dataSetsList.Clear()
        self.m_dataSetsList.AppendItems(ds_names)
    
    def OnDelDataSet(self, event):
        idx = self.m_dataSetsList.GetSelection()
        if idx<0:
            resp = wx.MessageBox('No list item selected', 'Operation cancelled', wx.OK | wx.ICON_WARNING)
            return            
        self.m_dataSets.pop(idx)
        ds_names = [i["name"] for i in self.m_dataSets]
        self.m_dataSetsList.Clear()
        self.m_dataSetsList.AppendItems(ds_names)

    
    def OnRefresh(self, event):
        self.m_extracted.SetLabel("")
        self.m_extracted.Refresh()
        self.m_extracted.Update()

        self.computeResults()
        self.LoadResults()


    def LoadResults(self): 
        input_file = resultsFolder1 + "/Phase1_results.txt"
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

    
    def OnNextPhase(self, event):
        output_file = resultsFolder1 + "/Phase1_excluded.txt"
        with open(output_file, 'w') as fp:
            for pg in self.m_excludedPages:
                print(pg, end='', file=fp) # The file-page strings contain a \n already...
                
        output_file = resultsFolder1 + "/Phase1_dataSets.json"
        jsonString = json.dumps(self.m_dataSets)
        jsonFile = open(output_file, "w")
        jsonFile.write(jsonString)
        jsonFile.close()        

        os.system('python VesselAI_extractData_2.py')

    
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
        if line in self.m_excludedPages:
            self.m_excludePage.SetValue(True)
        else:
            self.m_excludePage.SetValue(False)


app = wx.App(redirect=True)   # Error messages go to popup window
top = Frame("VesselAI_ExtractData_Phase_1")
top.Show()
app.MainLoop()
