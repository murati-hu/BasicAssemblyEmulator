REM   //N faktoriális
INP X  //N
GSB 14 
MOV F X //Eredeti szám eltárolása
INP Y 1 //Az utolsó szorzó beolvasása(1)
INP A 1 //Szorzat inicializálása
MLP A X //Összeg növelése
DEC X  //Visszaszámlálás
JMP 11 
GTO 7 
OUT A 
END  
REM   //Abszolut érték számító
INP Y 0
INP A -1
SIG 18 
MLP X A
RET  
