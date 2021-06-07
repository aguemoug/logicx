
export M4PATH="/data/dev/latex/circuit_macro"
m4 pgf.m4 code.m4 | dpic -g > code.tex
latex -interaction=nonstopmode loader.tex
dvisvgm -n1 -e loader.dvi -s >shape.svg
rm code.*
rm loader.*
