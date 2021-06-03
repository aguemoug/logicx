
    export M4PATH="/data/dev/latex/circuit_macro"
    m4 pgf.m4 code.m4 | dpic -g > code.tex
    latex -shell-escape -interaction=nonstopmode loader.tex > loader.out
    dvisvgm -n1 -e loader.dvi > loader.dat 2>&1
    rm code.*
    cp loader.svg shape.svg
    rm loader.*
    