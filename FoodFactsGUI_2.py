import Tkinter, sys
from xlrd import open_workbook


"""
Programmer: Jonathan Quass
Title: Food Facts GUI
Version 1.5

The purpose of this program is to aid a Food Facts employee in filtering
the list of ingredients before data entry. This is done by reading a master
excel file which contains a library of ingredients and alternate definitions.
Then, that library is used to check a potential list for any alternate
definitions. Upon finding them, the word is replaced with the prefered term
to ensure consistency in the data. If a term is not found as either a primary
term or alternate term, the term is output as an unknown.

2012-9-10
"""

### Defining the foodfacts object
class foodfacts_tk(Tkinter.Tk):

    
    def __init__(self, parent):
        Tkinter.Tk.__init__(self, parent)
        self.parent = parent
        self.initialize()


    ### Initializing object variables    
    sheet = []
    products = []
    alias = {}

    ### Initialize function
    def initialize(self):
        self.grid()

### BUTTONS

        #### Read Excel Button
        buttonExcel = Tkinter.Button(self, text = u"Read Excel File",
                                     command=self.OnButtonClickExcel)
        buttonExcel.grid(column = 2, row = 3, sticky='NEW')

        #### Get CSL Button
        buttonCSL = Tkinter.Button(self, text = u"Procss CSL",
                                   command = self.OnButtonClickCSL)
        buttonCSL.grid(column=1, row = 3, sticky='NEW')

### LABELS
        
        #### Status Bar
        self.labelVariable = Tkinter.StringVar()
        label = Tkinter.Label(self, textvariable=self.labelVariable,
                               fg = "yellow", bg = "black")
        label.grid(column=1, row = 0, columnspan=2, sticky='EW')
        self.labelVariable.set(u"Please Press Button To Process Excel File")

        ### Input Box Label
        self.labelInputVariable = Tkinter.StringVar()
        labelInput = Tkinter.Label(self, textvariable=self.labelInputVariable,
                                   fg = "black", bg = "gray")
        labelInput.grid(column = 0, row = 0, sticky='EW')
        self.labelInputVariable.set(u"Enter Comma Seperated List Here")

        ### Output Box Label
        self.labelOutputVariable = Tkinter.StringVar()
        labelOutput = Tkinter.Label(self, textvariable=self.labelOutputVariable,
                                    fg = "black", bg="gray")
        labelOutput.grid(column=0, row=4, sticky='EW')
        self.labelOutputVariable.set(u"Processed Comma Seperated List")

        ### Path Entry Label
        self.labelPathVariable = Tkinter.StringVar()
        labelPath = Tkinter.Label(self, textvariable=self.labelPathVariable,
                                  fg="black", bg="gray")
        labelPath.grid(column=1, row=1, columnspan = 2, sticky='SEW')
        self.labelPathVariable.set(u"Enter Excel File Path and Name Here")

        ### Unknown Input Label
        self.labelUnknownsVariable = Tkinter.StringVar()
        labelUnknowns = Tkinter.Label(self, textvariable=self.labelUnknownsVariable,
                                  fg="black", bg="gray")
        labelUnknowns.grid(column=1, row=4, columnspan = 2, sticky='EW')
        self.labelUnknownsVariable.set(u"Uknown Inputs")        

### ENTRY

        #### Path Input Box
        self.entryVariable = Tkinter.StringVar()
        self.entryCSL = Tkinter.Entry(self, textvariable=self.entryVariable)
        self.entryCSL.grid(column=1, row=2, sticky='NEW', columnspan = 2)
        self.entryVariable.set(u"C:\\Users\\sbutterfield\\Google Drive\\Ingredient Check\\foodfacts_masterlist.xls")        

### TEXT BOX

        #### CSL Input Box
        self.inputCSL = Tkinter.Text(self)
        self.inputCSL.grid(column = 0, row=1, rowspan = 3, sticky='EW')
        self.inputCSL.bind("<Return>", self.OnPressEnter)

        #### Corrected CSL Output Box
        self.outputCSL = Tkinter.Text(self)
        self.outputCSL.grid(column = 0, row=5, sticky='EW')

        #### Unknown Ingredients Output Box
        self.outputUnknowns = Tkinter.Text(self)
        self.outputUnknowns.grid(column=1, row=5, columnspan=2, sticky='EW')


### Grid Configuration Options
        self.grid_columnconfigure(0, weight=1)
        self.resizable(True, False)
        self.update()
        self.geometry(self.geometry())


### WIDGET BOUND FUNCTIONS

    ### Main Excel Function, reads the sheet and creates
    ### the master list of products and dictionary of all
    ### alternate terms
    def OnButtonClickExcel(self):
        self.sheet = self.getSheet()
        self.products = self.collectProducts(self.sheet)
        self.alias = self.createDictionary(self.sheet)
        if len(self.products) > 0:
            self.labelVariable.set(u"Library Built Without Error...")
            self.inputCSL.focus_set()
        else:
            self.labelVariable.set(u"Library empty. Please check file locations and names")

    ### Main Comma Seperate List Function, reads the CSL
    ### from the CSL input box, processes it and outputs
    ### the result into the CSL output box
    def OnButtonClickCSL(self):
        CSL = self.getCSL()
        unknownIngredients = self.processList(CSL, self.alias, self.products)
        self.outputCSL.delete(1.0, Tkinter.END)
        self.outputUnknowns.delete(1.0, Tkinter.END)
        if len(self.alias) == 0:
            self.labelVariable.set(u"Dictionary Empty: Please Check Excel File Name and Location")
            self.outputCSL.insert(Tkinter.END, "Default - C:\\Users\\sbutterfield\\Google Drive\\Ingredient Check\\foodfacts_masterlist.xls")
        else:
            self.labelVariable.set(u"Corrected List")
            self.outputCSL.insert(Tkinter.END, self.createString(CSL))
            self.outputUnknowns.insert(Tkinter.END, self.createString(unknownIngredients))
            self.outputCSL.focus_set()
      
    ### Binds Enter to CSL button click
    def OnPressEnter(self, other):
        self.OnButtonClickCSL()

### HELPER FUNCTIONS

    ### Retrieves the sheet using the user input location
    def getSheet(self):
        location = self.entryVariable.get()
        book = open_workbook(location)
        return book.sheet_by_index(0)

    ### Collects Products from column 0 from xls file
    def collectProducts(self, sheet):
        products = []
        for p in range(1, sheet.nrows):
            prod = unicode(sheet.cell(p, 0).value).lower()
            products.append(prod)
        return products

    ### Reads all alternate terms and attaches them to the
    ### prefered term
    def createDictionary(self, sheet):
        dictionary = {}                                   
        for x in range(1, sheet.nrows):                 
            for y in range(1, sheet.ncols):
                if sheet.cell(x, y).value != '':
                    entry = unicode(sheet.cell(x, y).value).lower()
                    definition = unicode(sheet.cell(x, 0).value).lower()
                    dictionary[entry] = definition
        return dictionary



    ### function retrieves text input in CSL text box and
    ### does preliminary filtering for all unwanted inputs
    def getCSL(self):
        CSL = self.inputCSL.get(1.0, Tkinter.END)
        CSL = CSL.lower()
        while CSL.find('\n') > -1:
            CSL = CSL[:CSL.find('\n')] + CSL[CSL.find('\n') + 1:]
        while CSL.find('\r') > -1:
            CSL = CSL[:CSL.find('\r')] + CSL[CSL.find('\r') + 1:]
        while CSL.find('  ') > -1:
            CSL = CSL[:CSL.find('  ')] + CSL[CSL.find('  ') + 1:]
        while CSL.find(' ,') > -1:
            CSL = CSL[:CSL.find(' ,')] + CSL[CSL.find(' ,')+ 1:]
        while CSL.find('.') > -1:
            CSL = CSL[:CSL.find('.')] + CSL[CSL.find('.')+ 1:]
        while CSL.find('[') > -1:
            CSL = CSL.replace('[', '(')
        while CSL.find(']') > -1:
            CSL = CSL.replace(']', ')')
        while CSL.find('{') > -1:
            CSL = CSL.replace('{', '(')
        while CSL.find('}') > -1:
            CSL = CSL.replace('}', ')')
        while CSL.find(', and') > -1:
            CSL = CSL.replace(', and', ',')
        CSL = CSL.rsplit(", ")
        return CSL

    
    def processList(self, CSL, alias, products):
        unknownList = []
        i = 0
        for l in CSL:
            if self.hasParen(l):
                filteredList = self.ignoreParenList(l)
                CSL[i] = self.checkAndReplace(l, filteredList, alias, products, unknownList)
            else:
                CSL[i] = self.checkAndReplace(l, [l], alias, products, unknownList)
            i = i + 1
        return unknownList

    def hasParen(self, word):
        if ')' in word or '(' in word:
            return True
        return False

    def ignoreParenList(self, word):
        while word.find('((') > -1:
            word = word[:word.find('((')] + word[word.find('((') + 1:]
        while word.find('))') > -1:
            word = word[:word.find('))')] + word[word.find('))') + 1:]
        while word.find(' (') > -1:
            word = word[:word.find(' (')] + word[word.find(' (') + 1:]
        while word.find(' )') > -1:
            word = word[:word.find(' )')] + word[word.find(' )') + 1:]
        wordCopy = word.rsplit('(')
        splitList = []
        for w in wordCopy:
            splitList.extend(w.rsplit(')'))
        return splitList

    def checkAndReplace(self, originalWord, listOfWords, dictionary, masterList, unknownWordList):
        for term in listOfWords:
            if term not in masterList:
                if term in dictionary:
                    originalWord = originalWord.replace(term, dictionary[term])
                else:
                    termOrder = self.checkWordOrderMaster(term, masterList)
                    if termOrder != "none":
                        if termOrder != '':
                            originalWord = originalWord.replace(term, termOrder)
                    else:
                        termOrder = self.checkWordOrderDictionary(term, dictionary)
                        if termOrder != "none":
                            if termOrder != '':
                                originalWord = originalWord.replace(term, termOrder)
                        else:
                            if term != '':
                                unknownWordList.append(term)
        return originalWord

    def checkWordOrderMaster(self, term, masterList):
        termWords = term.rsplit(' ')
        if len(termWords) > 1:
            for words in masterList:
                wordsSplit = words.rsplit(' ')
                if len(wordsSplit) == len(termWords):
                    found = True
                    for t in termWords:
                        if t not in wordsSplit:
                            found = False
                    if found:
                        return words
        return "none"

    def checkWordOrderDictionary(self, term, dictionary):
        termWords = term.rsplit(' ')
        if len(termWords) > 1:
            for key in dictionary:
                keySplit = key.rsplit(' ')
                if len(keySplit) == len(termWords):
                    found = True
                    for t in termWords:
                        if t not in keySplit:
                            found = False
                    if found:
                        return dictionary[key]
        return "none"
                
        
        

    def createString(self, list):
        if len(list) == 0:
            return "All Ingredients Known"
        else:
            return ', '.join(list)    
    

if __name__ == "__main__":
    app = foodfacts_tk(None)
    app.title('Food Facts')
    app.mainloop()
