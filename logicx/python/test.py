import uno
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.beans import PropertyValue
from com.sun.star.awt import Size

from pythonscript import ScriptContext
import os
CTX = uno.getComponentContext()
SM = CTX.getServiceManager()
dummy_code=r""" 
.PS
del = linewid
scale=25.25
Origin:Here
log_init
S:AND_gate at Origin +(35,10);
line from S.Out right ; "sp_$R$" ljust 

D1:Demux(4,DM1,L2) at Origin+(0,15);
D2:Demux(4,DM2,L2) at Origin-(0,15);

dot(at D1.In -(5,0),0.5,0);"sp_$1$" rjust; line to D1.In;
dot(at (D1.In,D1.Sel0) -(5,2),0.5,0);"sp_$L$" rjust; line to D1.Sel0-(0,2);
dot(at (D1.In,D1.Sel1) -(5,5),0.5,0);"sp_$M$" rjust; line to D1.Sel1-(0,5) then to D1.Sel1;
line right  from D1.Out3; line from Here to (Here,S.In1) then to S.In1;
line right    from D2.Out3; line from Here to (Here,S.In2) then to S.In2;
line from D1.In -(2.5,0) to D2.In -(2.5,0);
dot(at (D2.In,D2.Sel0) -(5,2),0.5,0);"sp_$H$" rjust; line to D2.Sel0-(0,2);
dot(at (D2.In,D2.Sel1) -(5,5),0.5,0);"sp_$P$" rjust; line to D2.Sel1-(0,5) then to D2.Sel1;

.PE
"""
latex_loader_code=r"""
\documentclass[tikz,convert={outfile=\jobname.svg}]{standalone}
%\usetikzlibrary{...}% tikz package already loaded by 'tikz' option
\begin{document}
\input{code.tex}
\end{document} 
"""
bash_script=r"""
export M4PATH="/data/dev/latex/circuit_macro"
m4 pgf.m4 code.m4 | dpic -g > code.tex
latex -interaction=nonstopmode loader.tex
dvisvgm -n1 -e loader.dvi > loader.dat 2>&1
#rm code.*
cp loader.svg shape.svg
#rm loader.*
"""





def write_text_to_file(text,file):
    code_file = open(file, "w")
    code_file.write(text)
    code_file.close()  





def loadMainGui(ctx):
    smgr = ctx.ServiceManager
    oDM = smgr.createInstance("com.sun.star.awt.UnoControlDialogModel")
    dialog_provider = smgr.createInstance("com.sun.star.awt.DialogProvider")
    dialog = dialog_provider.createDialog("vnd.sun.star.script:LogicxLibrary.main_dlg?location=application")
    txtbox = dialog.getControl("shapeCodeTectBox")
    txtbox.Text =dummy_code
    dialog_result =dialog.execute() 
    if(dialog_result == 1):
        print("drawing shape")
        shape_code = txtbox.Text
        print("code was :", shape_code)
        make_image(shape_code)



def create_dialog(ctx):
    loadMainGui(ctx)

def create_instance(name, with_context=False):
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance


class Logix:
    def __init__(self) -> None:
        self.imagePath="/data/dev/logix/shape.svg"
        ctxlocal = uno.getComponentContext()
        gslocal = ctxlocal.ServiceManager
        connecteur = gslocal.createInstanceWithContext( "com.sun.star.bridge.UnoUrlResolver", ctxlocal )
        self.CTX = connecteur.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
        self.SM = self.CTX.ServiceManager
        self.DESKTOP = self.SM.createInstanceWithContext( "com.sun.star.frame.Desktop",self.CTX)
        self.oDispatcher= SM.createInstanceWithContext("com.sun.star.frame.DispatchHelper", CTX)
        self.DOC =self.DESKTOP.getCurrentComponent()
        # Get the selected item
        self.DocCTL = self.DOC.getCurrentController()
   
    def copyImageToClipboard(self):
        img_url=uno.systemPathToFileUrl(self.imagePath)
        prop = PropertyValue()
        prop.Name ="Hidden"
        prop.Value = True
        oDrawDoc = self.DESKTOP.loadComponentFromURL( img_url, "_blank", 0,[prop])
        oDrawDocCtrl = oDrawDoc.getCurrentController()
        oDrawPage = oDrawDoc.getDrawPages().getByIndex(0)
        oImportShape = oDrawPage.getByIndex(0)
        oShapeSize = Size()
        #Copy the image to clipboard and close the draw document
        oDrawDocCtrl.select(oImportShape)    
        self.oDispatcher.executeDispatch( oDrawDocCtrl.Frame, ".uno:Copy", "", 0,[])
        oDrawDoc.close(True)
        oDrawDoc.dispose()
    def pastImageFromClipboard(self):
        #Paste image to the current document
        self.oDispatcher.executeDispatch( self.DocCTL.Frame, ".uno:Paste", "", 0,[])
    def buildImage(self,shape_code):
        write_text_to_file(shape_code,"code.m4")
        write_text_to_file(latex_loader_code,"loader.tex")
        if not os.path.exists('shape.sh'):
            write_text_to_file(bash_script,"shape.sh")
        os.system("bash shape.sh")





def main():
    plugin = Logix()
    plugin.buildImage(dummy_code)
    plugin.copyImageToClipboard()
    plugin.pastImageFromClipboard()
    #make_image(dummy_code)

main()


