REM   //N faktoriális
INP X  //N
MOV F X //Eredeti szám eltárolása
INP Y 1 //Az utolsó szorzó beolvasása(1)
INP A 1 //Szorzat inicializálása
MLP A X //Összeg növelése
DEC X  //Visszaszámlálás
JMP 10 
GTO 6 
OUT A 
END  
