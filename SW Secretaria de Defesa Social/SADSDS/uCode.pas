unit uCode;

interface

type
  matriz = array [0 .. 103, 0 .. 103] of double;   //matriz = array of array of double 25x25
  vetor = array [0 .. 103] of integer;
  arrayofarray = array of array of real;
  Pinteger = ^integer;
  P_mat = ^vetor;

  //Este procedimento recebe a matriz "a" resolve ela por Simplex e devolve a matriz "a" com a resposta.
 procedure simplx(var a: matriz; m, n, m1, m2, m3: integer; icase: Pinteger;
  izrov, iposv: vetor); cdecl; external 'SimplexDll.dll' name
  '?simplx@@YAXQAY0GI@NHHHHHPAH11@Z';


implementation

uses Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, ComCtrls, ExtCtrls, ComObj, Math;


end.
