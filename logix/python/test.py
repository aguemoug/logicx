# Copyright (C) 2021  Ghemougui abdessettar

# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.

# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.



from typing import Deque
import uno
import unohelper
import os

from com.sun.star.beans import PropertyValue
from com.sun.star.awt import Size
from com.sun.star.awt import XItemListener, XActionListener


latex_loader_code=r"""
\documentclass{standalone}
\usepackage{pgf,tikz}
\usepackage{color}
\begin{document}
\input{code.tex}
\end{document} 
"""
bash_script=r"""
export M4PATH="/data/dev/latex/circuit_macro"
m4 pgf.m4 code.m4 | dpic -g > code.tex
latex -interaction=nonstopmode loader.tex
dvisvgm -n1 -e loader.dvi -s >shape.svg
rm code.*
rm loader.*
"""




colors={
    "white":"FFFFFF",
    "black":"000000",
    "csorange":"FFBF91",
    "csbege":"A99B82",
    "mblack":"031716",
    "csyellow":"F2BB13",
    "cswhiteyellow":"DAF280",
    "csblue":"9EC4DB",
    "csgreenyellow":"B3D929"

}

#utility functions
def hex2rgb(hexstr):
     return tuple(int(hexstr[i:i+2], 16) for i in (0, 2, 4))
 

def text2file(text,file):
    code_file = open(file, "w")
    code_file.write(text)
    code_file.close()  

def loadDb():
    if not os.path.exists("data.lgx"):
        print("database file not found !")
        return None
    f = open('data.lgx','r')
    code=''
    cat=''
    subcat=''
    data ={}
    while True:
        line = f.readline()
        if not line: #end of file
            break
        if line.startswith('//') or line.isspace() :
            continue
        elif line.startswith("#;"): # end of a shape description
            (data[cat])[subcat]=code
            pass
        elif line.startswith("#$"): # start of a shape description 
            cat = line[2:].split('/')[0].strip()
            subcat = line[2:].split('/')[1].strip()
            code =''
            if not cat in data:
                data[cat]={}
        else:
            code += line +'\n'
                       
    f.close()
    return data


class CategorySelectedItemListener( unohelper.Base, XItemListener ):
    def __init__(self,logix_plugin):
        self.plugin = logix_plugin
    def itemStateChanged(self, event):
        self.plugin.subComboBox.removeItems(0,self.plugin.subComboBox.getItemCount())
        for i,x in enumerate( self.plugin.data[self.plugin.catComboBox.Text]):
            self.plugin.subComboBox.addItem(x,i)
        self.plugin.subComboBox.Text=''

class ShapeSelectedItemListener( unohelper.Base, XItemListener ):
    def __init__(self,logix_plugin):
        self.plugin = logix_plugin
    def itemStateChanged(self, event):
        self.plugin.shapeCodeTextBox.Text=self.plugin.data[self.plugin.catComboBox.Text][self.plugin.subComboBox.Text]

class Logix:
    def __init__(self,embedded) -> None:
        self.imagePath="/data/dev/logix/shape.svg"
        self.code=""
        self.setColor("black")
        if embedded:
            #if runing as a from libreoffice
            self.CTX = uno.getComponentContext()
            self.SM = self.CTX.ServiceManager
        else:
            # if runing as standalone script via uno bridge
            ctxlocal = uno.getComponentContext()
            gslocal = ctxlocal.ServiceManager
            connecteur = gslocal.createInstanceWithContext( "com.sun.star.bridge.UnoUrlResolver", ctxlocal )
            self.CTX = connecteur.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
            self.SM = self.CTX.ServiceManager
        
        self.DESKTOP = self.SM.createInstanceWithContext( "com.sun.star.frame.Desktop",self.CTX)
        self.oDispatcher= self.SM.createInstanceWithContext("com.sun.star.frame.DispatchHelper", self.CTX)
        self.DOC =self.DESKTOP.getCurrentComponent()
        self.DocCTL = self.DOC.getCurrentController()

        cfg = self.SM.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', self.CTX)
        prop = PropertyValue()
        prop.Name = 'nodepath'
        prop.Value = '/ooo.ext.logix.Registry/Settings'
        #self.CFG = cfg.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess', (prop,))

        self.data = loadDb()
    def loadUi(self):
        dialog_provider = self.SM.createInstance("com.sun.star.awt.DialogProvider")
        dialog = dialog_provider.createDialog("vnd.sun.star.script:LogicxLibrary.main_dlg?location=application")
        self.shapeCodeTextBox = dialog.getControl("shapeCodeTectBox")
        self.catComboBox = dialog.getControl("catComboBox")
        self.subComboBox = dialog.getControl("subComboBox")
        self.colorComboBox = dialog.getControl("colorComboBox")
        self.previewButton  = dialog.getControl("previewButton")
        #self.previewImageBox =dialog.getControl("previewImageBox")
        self.catComboBox.addItemListener(CategorySelectedItemListener(self))
        self.subComboBox.addItemListener(ShapeSelectedItemListener(self))       
        #self.previewButton.addActionListener(PreviewButtonListener(self))       
        #adding circuit categories 
        for i,cat in enumerate(self.data):
            self.catComboBox.addItem(cat,i+1)
        # adding colors 
        for i,cl in enumerate(colors):
            self.colorComboBox.addItem(cl,i+1)
        self.shapeCodeTextBox.Text =""
        dialog_result =dialog.execute() 
        if(dialog_result == 1):
            self.drawImage()
    
    def setColor(self,color):
        if color in colors:
            r,g,b = hex2rgb(colors[color])
            self.setRGB(r,g,b)
        else:
            self.setRGB(0,0,0)
        

    
    def drawImage(self):
        self.buildImage()
        self.copyImageToClipboard()
        self.pasteImageFromClipboard()
                 
    def setRGB(self,r,g,b):
        self.colorR=r/255.0
        self.colorG=g/255.0
        self.colorB=b/255.0
    def formatM4code(self,code):
        self.code =''
        self.code += ".PS\n"
        self.code += "setrgb({},{},{})\n".format(self.colorR,self.colorG,self.colorB)
        self.code += code + "\n"
        self.code += ".PE\n"
   
    def copyImageToClipboard(self):
        img_url=uno.systemPathToFileUrl(self.imagePath)
        prop = PropertyValue()
        prop.Name ="Hidden"
        prop.Value = True
        oDrawDoc = self.DESKTOP.loadComponentFromURL( img_url, "_blank", 0,[prop])
        oDrawDocCtrl = oDrawDoc.getCurrentController()
        try:
            oDrawPage = oDrawDoc.getDrawPages().getByIndex(0)
            oImportShape = oDrawPage.getByIndex(0)
            #Copy the image to clipboard and close the draw document
            oDrawDocCtrl.select(oImportShape)    
            self.oDispatcher.executeDispatch( oDrawDocCtrl.Frame, ".uno:Copy", "", 0,[])
            oDrawDoc.close(True)
            oDrawDoc.dispose()
        except:
            oDrawDoc.close(True)
            oDrawDoc.dispose()
    def pasteImageFromClipboard(self):
        #Paste image to the current document
        self.oDispatcher.executeDispatch( self.DocCTL.Frame, ".uno:Paste", "", 0,[])
    def buildImage(self):
        self.setColor(self.colorComboBox.Text)
        self.formatM4code(self.shapeCodeTextBox.Text)
        text2file(self.code,"code.m4")
        text2file(latex_loader_code,"loader.tex")
        if not os.path.exists('shape.sh'):
            text2file(bash_script,"shape.sh")
        os.system("bash shape.sh")



def main():   
    plugin = Logix(False)
    plugin.loadUi()

main()


