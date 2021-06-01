import uno
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
import os

def connect_to_oo():
    # Récupération du contexte d'exécution PyUNO
    ctxlocal = uno.getComponentContext()
    # Récupération du gestionnaire de services du contexte local PyUNO
    gslocal = ctxlocal.ServiceManager
    # création d'un connecteur logiciel UnoUrlResolver pour dialoguer en appel distant de procédures RPC
    connecteur = gslocal.createInstanceWithContext( "com.sun.star.bridge.UnoUrlResolver", ctxlocal )
    # Récupération du contexte de LibreOffice
    ctx = connecteur.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
    # Récupération du gestionnaire de services du contexte de Libreoffice
    gestionaireadministration = ctx.ServiceManager
    # Récupérer l'objet principal du bureau LibreOffice
    bureau = gestionaireadministration.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    # Récupérer l'affichage actuel
    cadre = bureau.getCurrentFrame()
    # Créer la fenêtre de la boite de dialogue 
    fenêtre = cadre.getContainerWindow()

    # Récupérer les outils de gestion de la fenêtre
    boîteàoutils = fenêtre.getToolkit()

    # Affectation du titre de la boîte de dialogue à une variable
    titre = "Python"

    # Affectution du message à afficher dans la boîte de dialogue à une variable
    message = "Bonjour à tous !"

    # Création du modèle de boite de dialogue
    boîtemessage = boîteàoutils.createMessageBox(fenêtre, MESSAGEBOX, 1, titre, message)

    # Création de l'objet boîte de dialogue
    boîtemessage.execute()

dummy_code=r""" 
.PS
include(pstricks.m4)
log_init
S:Mux(2,Mux 2x1,S1);

.PE
"""
latex_loader_code=r"""
\documentclass[tikz,convert={outfile=\jobname.svg}]{standalone}
%\usetikzlibrary{...}% tikz package already loaded by 'tikz' option
\begin{document}
\input{code.tex}
\end{document} 
"""

def write_text_to_file(text,file):
    code_file = open(file, "w")
    code_file.write(text)
    code_file.close()  
def clear():
    files=["code.tex","code.pic","code.svg","code.m4","loader.fls","loader.tex","loader.synctex.gz","loader.log","loader.dvi","loader.aux","loader.fdb_latexmk"]
    for f in files :
        try:
            os.remove(f)
        except:
            pass
def compile_image():
    os.system("latex -interaction=nonstopmode loader.tex")
    os.system("dvisvgm --no-fonts loader.dvi image.svg")
def make_image(shape_code):
    write_text_to_file(shape_code,"code.m4")
    write_text_to_file(latex_loader_code,"loader.tex")
    os.system('m4 code.m4 | dpic -g > code.tex')
    #compile_image()
    
    #clear()

def main():
    make_image(dummy_code)

main()


