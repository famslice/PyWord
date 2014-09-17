# By : Haris Ahmad and Mudassir Chugtahi
# PyWord

import bisect
import wx, os, sys, cStringIO
import wx.lib.agw.ribbon as RB
from wx.html import HtmlEasyPrinting
import wx.lib.agw.rulerctrl as RC
from wx.richtext import *

class PyWindow(wx.Frame):

                                                                # ID = 800 OR GREATER
    LETTER_KEY_CODES = [65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78,
                        79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 97, 89,
                        99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109,
                        110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120,
                        121, 122]
    
    word_suggest_enabled = False# Enable Word Sugest
    spell_check_enabled = False #Enable Spell Check
    
    # Fetch words from dictionary file to initialize dictionary
    dictionary = open("dictionary.txt").read().split()
    dictionary.sort()
    
    suggest_dictionary = open("suggest_dict.txt").read().split()
    suggest_dictionary.sort() # Sort in alphabetical order

    def __init__(self,parent,id=-1,title="",pos=wx.DefaultPosition,size=wx.DefaultSize,style=wx.DEFAULT_FRAME_STYLE):

        self.frame = wx.Frame.__init__(self,parent,id,title,pos,size,style) # Initiate parent window
        
        self.SetIcon(wx.Icon("Icons/Pie.ico", wx.BITMAP_TYPE_ICO)) # Set window icon

        self.ribbon = RB.RibbonBar(self, -1) # Initiate Ribbon Bar
        self.CreateStatusBar() # Initiate Status Bar

        ### ------------------------------------------- Home Tab: id = 100 -------------------------------------------- ###
        home = RB.RibbonPage(self.ribbon, -1, "Home") # Home tab on bar

        filePanel = RB.RibbonPanel(home, -1, "File") # File menu 
        fileBar = RB.RibbonButtonBar(filePanel)        
        fileBar.AddSimpleButton(101, "New",wx.Bitmap("Icons/stock_new.png", wx.BITMAP_TYPE_PNG),"") # New document
        fileBar.AddSimpleButton(102, "Open",wx.Bitmap("Icons/stock_open.png", wx.BITMAP_TYPE_PNG),"") # Open file tool
        fileBar.AddHybridButton(103, "Save",wx.Bitmap("Icons/stock_save.png", wx.BITMAP_TYPE_PNG)) # Save file tool
        fileBar.AddHybridButton(104,"Print", wx.Bitmap("Icons/stock_print.png", wx.BITMAP_TYPE_PNG),"") # Print

        alignmentPanel = RB.RibbonPanel(home,-1, "Alignment") # Alignment menu 
        alignBar = RB.RibbonButtonBar(alignmentPanel)
        alignBar.AddSimpleButton(105, "Left", wx.Bitmap("Icons/align_left.png", wx.BITMAP_TYPE_PNG),"") # Left Align
        alignBar.AddSimpleButton(106, "Center", wx.Bitmap("Icons/align_center.png", wx.BITMAP_TYPE_PNG),"") # Centre Align
        alignBar.AddSimpleButton(107, "Right", wx.Bitmap("Icons/align_right.png", wx.BITMAP_TYPE_PNG),"") # Right Align
        
        toolsPanel = RB.RibbonPanel(home, -1 , "Text") # Text menu
        toolsBar =  RB.RibbonButtonBar(toolsPanel)        
        toolsBar.AddSimpleButton(108,"Standard", wx.Bitmap("Icons/bullet_point.png", wx.BITMAP_TYPE_PNG),"") # Standard Bullets
        toolsBar.AddSimpleButton(109,"Number", wx.Bitmap("Icons/numb_point.png", wx.BITMAP_TYPE_PNG),"") # Numbered Bullets
        toolsBar.AddSimpleButton(110,"Bold" , wx.Bitmap("Icons/stock_text_bold.png", wx.BITMAP_TYPE_PNG),"") # Bold text
        toolsBar.AddSimpleButton(111,"Italic" , wx.Bitmap("Icons/stock_text_italic.png", wx.BITMAP_TYPE_PNG),"") # Italic text
        toolsBar.AddSimpleButton(112,"Underline", wx.Bitmap("Icons/stock_text_underline.png", wx.BITMAP_TYPE_PNG),"") # Underline Text
        toolsBar.AddSimpleButton(119,"Highlight",wx.Bitmap("Icons/highlighter_icon.png",wx.BITMAP_TYPE_PNG),"") # Highlight text
        
        lineSpace = RB.RibbonPanel(home, -1, "Spacing") # Spacing menu
        lineBar = RB.RibbonButtonBar(lineSpace)
        lineBar.AddSimpleButton(113,"Normal", wx.Bitmap("Icons/spacing_normal.png", wx.BITMAP_TYPE_PNG),"") # Single line spacing
        lineBar.AddSimpleButton(114,"Half", wx.Bitmap("Icons/spacing_half.png", wx.BITMAP_TYPE_PNG),"") # 1 1/2 line spacing
        lineBar.AddSimpleButton(115,"Double", wx.Bitmap("Icons/spacing_double.png", wx.BITMAP_TYPE_PNG),"") # Double line spacing

        tabPanel = RB.RibbonPanel(home,-1, "Indent") # Indent menu
        tabBar = RB.RibbonButtonBar(tabPanel)
        tabBar.AddSimpleButton(116,"Right", wx.Bitmap("Icons/arrow_right.png", wx.BITMAP_TYPE_PNG),"") # Right indent
        tabBar.AddSimpleButton(117,"Left", wx.Bitmap("Icons/arrow_left.png", wx.BITMAP_TYPE_PNG),"") # Left indent
        
        picPanel = RB.RibbonPanel(home, -1, "Add Pictures") # Pictures menu
        picBar = RB.RibbonButtonBar(picPanel)    
        picBar.AddSimpleButton(118,"Import", wx.Bitmap("Icons/insert_image.png", wx.BITMAP_TYPE_PNG),"") # Import pics      

        ### ------------------------------------------- Font Tab: id = 200 -------------------------------------------- ###
        font_tab = RB.RibbonPage(self.ribbon, -1, "Font") # Font tab created
        fontPanel = RB.RibbonPanel(font_tab, 400, "Font") # Font gallery
        self.addColours(fontPanel, 200)
        fontPanel2 = RB.RibbonPanel(font_tab, -1, "Fonts") # Fonts bar
        fonts = RB.RibbonButtonBar(fontPanel2)
        fonts.AddSimpleButton(201,"Font", wx.Bitmap("Icons/edit_pic.jpg", wx.BITMAP_TYPE_JPEG),"")
        fonts.AddSimpleButton(202,"Words", wx.Bitmap("Icons/suggestion_icon.png", wx.BITMAP_TYPE_PNG),"")
        fonts.AddSimpleButton(203,"Check", wx.Bitmap("Icons/suggestion_icon.png", wx.BITMAP_TYPE_PNG),"")
        fontPanel.Realize()
        fontPanel2.Realize()

        ### ------------------------------------------- End Ribbon ---------------------------------------------------- ###

        self.rtc = RichTextCtrl(self, style=wx.HSCROLL|wx.NO_BORDER) # Init rich text box
        wx.CallAfter(self.rtc.SetFocus) # Display text box
        self.rtc.BeginFont(wx.Font(15, wx.SWISS , wx.NORMAL, wx.NORMAL, False, "Calibri"))

        box = wx.BoxSizer(wx.VERTICAL) # Resizes widgets within window when resized

        box.Add(self.ribbon,0,wx.EXPAND) # Resizes ribbon appropriately
        box.Add(self.rtc,1,wx.EXPAND) # Fills in remaining space with text box

        self.ribbon.Realize() # Displays ribbon
        self.SetSizer(box) # Initiate
        self.Show(True)
        
        self.bindEvents()
            
    def bindEvents(self):
        # Bind buttons to respective handlers
        
        self.Bind(wx.EVT_CLOSE, self.exitWindow, id=wx.ID_ANY)
        
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.newWindow, id=101)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.openFile, id=102)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.saveFile, id=103)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_DROPDOWN_CLICKED, self.saveDropdown, id=103)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.printDoc, id=104)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_DROPDOWN_CLICKED, self.printDropdown, id=104)
        
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.font, id=201)

        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.alignLeft, id=105)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.alignCentre, id=106)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.alignRight, id=107)
        
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.BulletPoint, id=108)
        
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.ApplyBold, id=110)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.ApplyItalic, id=111)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.ApplyUnderline, id=112)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.Highlight, id= 119)
        
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.LineSpacingNormal, id=113)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.LineSpacingHalf, id=114)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.LineSpacingDouble, id=115)
        
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.IndentRight, id = 116)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.IndentLeft, id = 117)
        
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.InsertPicture, id=118)
        
        self.Bind(RB.EVT_RIBBONGALLERY_SELECTED, self.colourSet, id=200)

        #------------------------------------------------------------------------
        self.Bind(wx.EVT_KEY_UP, self.onKeyPressed)#run the event everytime a key goes up
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.WordSuggestON, id = 202)
        self.Bind(RB.EVT_RIBBONBUTTONBAR_CLICKED, self.SpellCheckON, id = 203)

    def IndentRight(self,event):
        attr = TextAttrEx() # Initiate text attribute control
        attr.SetFlags(TEXT_ATTR_LEFT_INDENT)
        ip = self.rtc.GetInsertionPoint()
        if self.rtc.GetStyle(ip, attr):
            r = RichTextRange(ip, ip) # Set range of selection
            if self.rtc.HasSelection():
                r = self.rtc.GetSelectionRange()

            attr.SetLeftIndent(attr.GetLeftIndent() + 100)
            attr.SetFlags(TEXT_ATTR_LEFT_INDENT) # Set indent
            self.rtc.SetStyle(r, attr)
            
    def IndentLeft(self,event):
        attr = TextAttrEx()
        attr.SetFlags(TEXT_ATTR_LEFT_INDENT)
        ip = self.rtc.GetInsertionPoint()
        if self.rtc.GetStyle(ip, attr):
            r = RichTextRange(ip, ip)
            if self.rtc.HasSelection():
                r = self.rtc.GetSelectionRange()

        if attr.GetLeftIndent() >= 100:
            attr.SetLeftIndent(attr.GetLeftIndent() - 100)
            attr.SetFlags(TEXT_ATTR_LEFT_INDENT)
            self.rtc.SetStyle(r, attr)

    def font(self,event):
        default_font = wx.Font(12, wx.SWISS , wx.NORMAL, wx.NORMAL, False, "Calibri")
        data = wx.FontData()
        if sys.platform == 'win32':
            data.EnableEffects(True)
        data.SetInitialFont(default_font)
        dlg = wx.FontDialog(self, data)
        if dlg.ShowModal() == wx.ID_OK:
            data = dlg.GetFontData()
            font = data.GetChosenFont()
            self.rtc.BeginFont(font)
            self.rtc.BeginFontSize(font.GetPointSize())
            colour = data.GetColour()
            if self.rtc.HasSelection(): # If there is a selection, then change the selected text's font
                selection = self.rtc.GetSelectionRange()
                atr = TextAttrEx()
                atr.SetTextColour(colour)
                atr.SetFontFaceName(str(font))
                self.rtc.SetStyle(selection,atr)                
            else: # If not then start new text
                self.rtc.BeginTextColour(colour)
                text = 'Face: %s, Size: %d, Colour: %s' % (font.GetFaceName(), font.GetPointSize(),  colour.Get())
                self.SetStatusText(text)
        dlg.Destroy()
        
    def addColours(self, panel, id):
        self.colours = RB.RibbonGallery(panel, id, name = "")
        black = self.colours.Append(wx.Bitmap("Colours/black.bmp", wx.BITMAP_TYPE_BMP), id) # Image
        self.colours.SetItemClientData(black,(0,0,0)) # Data
        red = self.colours.Append(wx.Bitmap("Colours/red.bmp", wx.BITMAP_TYPE_BMP), id)
        self.colours.SetItemClientData(red,(255,0,0))
        blue = self.colours.Append(wx.Bitmap("Colours/blue.bmp", wx.BITMAP_TYPE_BMP), id)
        self.colours.SetItemClientData(blue,(0,0,255))
        green = self.colours.Append(wx.Bitmap("Colours/green.bmp", wx.BITMAP_TYPE_BMP), id)
        self.colours.SetItemClientData(green,(0,255,0))
        white = self.colours.Append(wx.Bitmap("Colours/white.bmp", wx.BITMAP_TYPE_BMP), id)
        self.colours.SetItemClientData(white,(255,255,255))
        white = self.colours.Append(wx.Bitmap("Colours/white.bmp", wx.BITMAP_TYPE_BMP), id)
        self.colours.SetItemClientData(white,(255,255,255))
        self.colours.SetSelection(black)
        self.colours.Realize()
        
    def colourSet(self,event):
        colour = self.colours.GetItemClientData(event.GetGalleryItem()) # Reads button for data (which colour)
        self.textColour = colour
        self.rtc.BeginTextColour(colour)

    def Highlight(self,event):
        attr =  TextAttrEx()
        if self.rtc.HasSelection():
            rang = self.rtc.GetSelectionRange()
            attr.SetBackgroundColour((225,225,0))
            self.rtc.SetStyle(rang, attr)
            
    def InsertPicture(self, event):
        dlg = wx.FileDialog(self, "Choose a file",  os.getcwd(), "", "PNG files (*.png)|*.png|All files (*.*)|*|", wx.OPEN|wx.CHANGE_DIR)                                                        
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
        self.rtc.WriteImageFile(path, wx.BITMAP_TYPE_ANY) # Insert selected image      
        
    def BulletPoint(self,event):
        self.rtc.BeginStandardBullet('standard/circle',100,60)
        self.rtc.WriteText("")
        self.rtc.EndSymbolBullet()

        
    def LineSpacingNormal (self, event): #line Spacing , 20 = double, 15 = 1.5 , 10 = normal
        attr = TextAttrEx()
        attr.SetFlags(TEXT_ATTR_LINE_SPACING)
        ip = self.rtc.GetInsertionPoint()
        if self.rtc.GetStyle(ip, attr):
            r = RichTextRange(ip, ip)
            if self.rtc.HasSelection():
                r = self.rtc.GetSelectionRange()
            attr.SetFlags(TEXT_ATTR_LINE_SPACING)
            attr.SetLineSpacing(10)
            self.rtc.SetStyle(r, attr)
        
    def LineSpacingHalf (self, event): #line Spacing , 20 = double, 15 = 1.5 , 10 = normal
        attr = TextAttrEx()
        attr.SetFlags(TEXT_ATTR_LINE_SPACING)
        ip = self.rtc.GetInsertionPoint()
        if self.rtc.GetStyle(ip, attr):
            r = RichTextRange(ip, ip)
            if self.rtc.HasSelection():
                r = self.rtc.GetSelectionRange()
            attr.SetFlags(TEXT_ATTR_LINE_SPACING)
            attr.SetLineSpacing(15)
            self.rtc.SetStyle(r, attr)

    def LineSpacingDouble (self, event): #line Spacing , 20 = double, 15 = 1.5 , 10 = normal
        attr = TextAttrEx()
        attr.SetFlags(TEXT_ATTR_LINE_SPACING)
        ip = self.rtc.GetInsertionPoint()
        if self.rtc.GetStyle(ip, attr):
            r = RichTextRange(ip, ip)
            if self.rtc.HasSelection():
                r = self.rtc.GetSelectionRange()
            attr.SetFlags(TEXT_ATTR_LINE_SPACING)
            attr.SetLineSpacing(20)
            self.rtc.SetStyle(r, attr)
        
    def ApplyUnderline(self,event):
        self.rtc.ApplyUnderlineToSelection()

    def ApplyBold (self,event):
        self.rtc.ApplyBoldToSelection()

    def ApplyItalic(self,event):
        self.rtc.ApplyItalicToSelection()

    def alignLeft(self,event):
        self.rtc.ApplyAlignmentToSelection(wx.TEXT_ALIGNMENT_LEFT)

    def alignCentre(self,event):
        self.rtc.ApplyAlignmentToSelection(wx.TEXT_ALIGNMENT_CENTER)

    def alignRight(self,event):
        self.rtc.ApplyAlignmentToSelection(wx.TEXT_ALIGNMENT_RIGHT)

    def newWindow(self,event):
        PyWindow(None, -1, "PyWord", size=(1024,768))

    def openFile(self,event):
        dlg = wx.FileDialog(self, "Choose a file", os.getcwd(), "", "XML files (*.xml)|*.xml|All files (*.*)|*|", wx.OPEN|wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            name = os.path.basename(path)
            dlg.Destroy()

            handler = RichTextXMLHandler()
            handler.SetFlags(RICHTEXT_HANDLER_INCLUDE_STYLESHEET)
            handler.LoadFile(self.rtc.GetBuffer(),name)
            self.rtc.Refresh()          

    def saveFile(self,event):
        self.saveFileAs(event)

    def saveDropdown(self,event):
        menu = wx.Menu()
        menu.Append(101, "Save")
        menu.Append(102, "Save As")
        self.Bind(wx.EVT_MENU, self.saveFile, id=101)
        self.Bind(wx.EVT_MENU, self.saveFileAs, id=102)
        event.PopupMenu(menu)

    def saveFileAs(self,event):
        dlg = wx.FileDialog(self, "Save file as...", os.getcwd(), "", "XML files (*.xml)|*.xml|All files (*.*)|*|", wx.SAVE|wx.OVERWRITE_PROMPT|wx.CHANGE_DIR)
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            name = os.path.basename(path)
            dlg.Destroy()
            
            doc = open(str(name), "w")
            buf = cStringIO.StringIO()
            handler = RichTextXMLHandler()
            handler.SetFlags(RICHTEXT_HANDLER_INCLUDE_STYLESHEET)
            handler.SaveStream(self.rtc.GetBuffer(), buf)
            doc.write(buf.getvalue())
            doc.close()

    def printDropdown(self,event):
        menu = wx.Menu()
        menu.Append(101, "Print")
        menu.Append(102, "Print Preview")
        self.Bind(wx.EVT_MENU, self.printDoc, id=101)
        self.Bind(wx.EVT_MENU, self.printPreview, id=102)
        event.PopupMenu(menu)

    def printPreview(self,event):
        printout1 = RichTextPrintout()
        printout1.SetRichTextBuffer(self.rtc.GetBuffer())   # Set the rich text buffer
        printout2 = RichTextPrintout()
        printout2.SetRichTextBuffer(self.rtc.GetBuffer())
        data = wx.PrintDialogData()
        data.SetAllPages(True)  
        data.SetCollate(True)   
        datapr = wx.PrintData() 
        data.SetPrintData(datapr)
        # Impression
        preview = wx.PrintPreview(printout1, printout2, data)
        if not preview.Ok():
            return
        pfrm = wx.PreviewFrame(preview, self, "Print Preview Test")
        pfrm.Initialize()
        pfrm.SetPosition(self.GetPosition())
        pfrm.SetSize(self.GetSize())
        pfrm.Show(True) 
        

    def printDoc(self,event):        
        """ Handle the Print menu item """
        # Create a RichTextPrinting helper object
        rtp = RichTextPrinting("Print test")
        # Call the Page Setup Dialog
        #rtp.PageSetup()
        # Calling Print Preview crashes and burns for me with wxPython 2.8.10.1 on Windows
        # and wxPython 2.8.9.2 on OS X.
        # rtp.PreviewBuffer(self.txtCtrl.GetBuffer())
        # Send the output to the printer
        rtp.PrintBuffer(self.rtc.GetBuffer())

    def exitWindow(self,event):
        dlg = wx.MessageDialog(self, "Save before Exit?", "Confirm Exit", wx.YES_NO | wx.YES_DEFAULT | wx.CANCEL | wx.ICON_QUESTION)
        var = dlg.ShowModal()
        if var == wx.ID_YES:
            self.saveFile(event) # Save then close
            self.Destroy()
        elif var == wx.ID_CANCEL: dlg.Destroy()
        else: self.Destroy() # Close otherwise

    ###++++++++++++++++++++++++++++++++++++++++++++++++++++++===WORD SOGGEST AND SPELLCHECK+++++++++++++++++++++++++++++++++++++++
##                                                            # ID = 800 OR GREATER
##    LETTER_KEY_CODES = [65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78,
##                        79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 97, 89,
##                        99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109,
##                        110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120,
##                        121, 122]
##    
##    word_suggest_enabled = False# Enable Word Sugest
##    spell_check_enabled = False #Enable Spell Check
##    
##    # Fetch words from dictionary file to initialize dictionary
##    dictionary = open("dictionary.txt").read().split()
##    dictionary.sort()
##    
##    suggest_dictionary = open("suggest_dict.txt").read().split()
##    suggest_dictionary.sort() # Sort in alphabetical order

    def SpellCheckON(self,event):
        if self.spell_check_enabled == False:
            self.spell_check_enabled = True
        else:
            self.spell_check_enabled = False

    def WordSuggestON(self,event):
        if self.word_suggest_enabled == False:
            self.word_suggest_enabled = True
        else:
            self.word_suggest_enabled = False

    def onKeyPressed(self, event):
        """
        Event handler for key presses
        """
        
        if event.GetKeyCode() in self.LETTER_KEY_CODES: # Check if the keys are in the list
            current_pos = self.rtc.GetInsertionPoint()
            word_start_pos = self.getStartOfWord(self.rtc.FindNextWordPosition(-1))
            if self.word_suggest_enabled: # Check if word suggest is enabled
                self.wordSuggest(current_pos, word_start_pos)
                self.range = self.rtc.GetSelectionRange()
            if self.spell_check_enabled:    #check if Spell Check is Enabled
                current_word = self.rtc.GetRange(word_start_pos, current_pos)
                spell_check_style = TextAttrEx()
                spell_check_style.SetTextColour((204, 0, 0))
                spell_check_range = RichTextRange(word_start_pos, current_pos)
                orig_style = self.rtc.GetDefaultStyle() #Get the defult style using so it can be changed ot that is the wor dis correct
                orig_style.SetTextColour((0, 0, 0))
                if self.isWordInDictionary(current_word) == False:
                    self.rtc.SetStyle(spell_check_range, spell_check_style)
                else:
                    print orig_style.GetTextColour()
                    self.rtc.SetStyle(spell_check_range, orig_style)
        
        if event.GetKeyCode() == wx.WXK_RIGHT:
            if (self.word_suggest_enabled == True) and (self.range.GetEnd() != self.range.GetStart()):
                self.rtc.SetInsertionPoint(self.range.GetEnd())
    
    def wordSuggest(self, current_pos, word_start_pos):
        #Suggest The Actual Word
        
        current_word = self.rtc.GetRange(word_start_pos, current_pos) # take the range of begning of the word to the end
        if len(current_word) > 3: # Only suggest word if it is longer than three letters
            suggested_index = self.indexOf(current_word)# Check  current word in the dictonry
            print suggested_index
            if suggested_index == -1:
                print "Suggestion for \"", current_word, "\" is \"None\""
            else:
                print "Suggestion for \"", current_word, "\" is \"", self.suggest_dictionary[suggested_index], "\""
                print (current_pos - word_start_pos), " ", len(self.suggest_dictionary[suggested_index])
                if len(self.suggest_dictionary[suggested_index]) > (current_pos - word_start_pos):
                    self.rtc.WriteText(self.suggest_dictionary[suggested_index][(current_pos - word_start_pos):])#suggest the letter than have not yet been typed
                    self.rtc.SetSelection(current_pos, (word_start_pos + len(self.suggest_dictionary[suggested_index])))
                range = self.rtc.GetSelectionRange()
                if range.GetEnd() != range.GetStart():
                    self.rtc.MoveCaret(range.GetStart() - 1)
    
    def onNextWordClicked(self, event):
        """
        Event handler for move cursor to next word button (id = 101)
        """
        
        print "Position found: ", self.rtc.FindNextWordPosition(1)
        self.rtc.SetInsertionPoint(self.rtc.FindNextWordPosition(1))
    
    def onPreviousWordClicked(self, event):
        """
        Event handler for move cursor to previous word button (id = 102)
        """
        
        print "Position found: ", self.rtc.FindNextWordPosition(-1)
        self.rtc.SetInsertionPoint(self.rtc.FindNextWordPosition(-1))
    
    def predictText(self, key_code):
        print "Key pressed: ", self.toChar(key_code)

    def toChar(self, key_code):
        #Check What key is pressed
        if 97 <= key_code <= 122:
            return chr(key_code)
        else:
            return chr(key_code).lower()
    
    def getStartOfWord(self, previous_word_pos):
        #Get the Position at the Start of the word
        if previous_word_pos == -1:
            return 0
        else:
            while ord(self.rtc.GetValue()[previous_word_pos]) not in self.LETTER_KEY_CODES:
                previous_word_pos += 1
            return previous_word_pos
    
    def getEndOfWord(self, next_word_pos):
        #get the position at the End of the Word
        if ord(self.rtc.GetValue()[next_word_pos]) in self.LETTER_KEY_CODES:
            return next_word_pos + 1
        else:
            return next_word_pos
    
    def indexOf(self, word):
        for i in range(len(self.suggest_dictionary)):
            if self.suggest_dictionary[i].startswith(word.lower()):
                return i
        return -1
    
    def isWordInDictionary(self, word):
        # Check if the wor dis in the Dictonry or not
        i = bisect.bisect_left(self.dictionary, word.lower())
        if i != len(self.dictionary) and self.dictionary[i] == word.lower():
            return True
        return False

    ###+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

if __name__ == "__main__":

    app = wx.App(0)

    try:

        frame = PyWindow(None, -1, "PyWord", size=(1024,768))

        app.MainLoop()

    finally:

        del app
