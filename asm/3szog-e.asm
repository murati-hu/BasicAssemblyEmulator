REM   //Háromszög oldalai
LET D 24 //Tételezzük fel, hogy rossz, tehát D=0
INP A  //A oldal
INP B  //B oldal
INP C  //C oldal
REM   //a<b+c ???
MOV X A
MOV Y B
ADD Y C
SIG 22  //Megbukott?
REM   //b<a+c ???
MOV X B
MOV Y A
ADD Y C
SIG 22 
REM   //c<a+b ???
MOV X C
MOV Y A
ADD Y B
SIG 22 
INC D 
OUT D 
END  
STR 0 
