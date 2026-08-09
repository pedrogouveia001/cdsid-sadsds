unit uHasseDiagram;

interface

uses
  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, SimpleGraph
{$IFDEF COMPILER7_UP}, XPMan {$ENDIF},
  IWCompLabel, IWCompButton, jpeg, IWVCLBaseControl, IWBaseControl,
  IWBaseHTMLControl, IWControl, IWExtCtrls, Controls, IWVCLBaseContainer,
  IWContainer, IWHTMLContainer, IWHTML40Container, IWRegion, Menus, ActnList,
  ImgList, StdCtrls, ComCtrls, ToolWin, Buttons,
  ExtCtrls, CheckLst, messages, Dialogs, Windows, Variants,
  Graphics, Forms, ExtDlgs;

type
  TElasticNode = class(TEllipticNode)
  private
  public
    constructor Create(AOwner: TSimpleGraph; cX, cY: Integer); reintroduce;
  end;

  TElasticEdge = class(TGraphLink)
  public
    constructor Create(AOwner: TSimpleGraph; SourceNode,
      TargetNode: TElasticNode); reintroduce;
    constructor Create1(AOwner: TSimpleGraph; SourceNode,
      TargetNode: TElasticNode); reintroduce;
    constructor Create2(AOwner: TSimpleGraph; SourceNode,
      TargetNode: TElasticNode); reintroduce;
    constructor Create3(AOwner: TSimpleGraph; SourceNode,
      TargetNode: TElasticNode); reintroduce;
  end;

type
  TfrmhasseD = class(TIWAppForm)
    rgnHasseDiagram: TIWRegion;
    IWImage15: TIWImage;
    btnclose: TIWButton;
    lblhasse: TIWLabel;
    IWButton1: TIWButton;
    IWButton2: TIWButton;
    imghasse: TIWImage;
    procedure IWButton1Click(Sender: TObject);
    procedure AMD;
    procedure QuebraPalavras(Texto: String);
    function SplitString(CountWords: array of String; Texto: string): Integer;
    procedure DomAlt2(SimpleGraph: TSimpleGraph; Node: Array of TElasticNode);
    procedure IWButton2Click(Sender: TObject);
    procedure rgnHasseDiagramCreate(Sender: TObject); // limpa da linha da alt analisada na matriz Node5 as alt que
    // não terão ligações com a mesma, devido a transitiviade
  public
  end;

implementation

uses UserSessionUnit;
{$R *.dfm}

// define tamanho e cor dos nós
constructor TElasticNode.Create(AOwner: TSimpleGraph; cX, cY: Integer);
begin
  inherited Create(AOwner);
  SetBounds(cX + Owner.Width div 2 - 10, cY + Owner.Height div 2 - 10, 100, 70);
  NodeOptions := NodeOptions - [gnoResizable];
  Brush.Color := clwebwhite; // cor do nó branco
end;

constructor TElasticEdge.Create1(AOwner: TSimpleGraph;
  SourceNode, TargetNode: TElasticNode);
begin
  inherited CreateNew(AOwner, SourceNode, [], TargetNode);
  Options := Options - [goSelectable];
  Brush.Color := clwebRed;
  Pen.Color := clwebRed;
  BeginStyle := lsDiamond; // tipo da geometria do inicio da seta
  EndStyle := lsArrow; // tipo da geometria do fim da seta
  BeginSize := 1; // tamanho da geometria do inicio da seta
  EndSize := 3; // tamanho da geometris do fim da seta
end;

{ TElasticEdge }
// propriedades das setas que ligam os nós quando há dominância
constructor TElasticEdge.Create(AOwner: TSimpleGraph;
  SourceNode, TargetNode: TElasticNode);
begin
  inherited CreateNew(AOwner, SourceNode, [], TargetNode);
  Options := Options - [goSelectable];
  Brush.Color := clwebblack;
  BeginStyle := lsDiamond; // tipo da geometria do inicio da seta
  EndStyle := lsArrow; // tipo da geometria do fim da seta
  BeginSize := 1; // tamanho da geometria do inicio da seta
  EndSize := 3; // tamanho da geometris do fim da seta
end;

// propriedades das setas que ligam os nós quando há indiferença
constructor TElasticEdge.Create2(AOwner: TSimpleGraph;
  SourceNode, TargetNode: TElasticNode);
begin
  inherited CreateNew(AOwner, SourceNode, [], TargetNode);
  Options := Options - [goSelectable];
  Pen.Color := clWebGRAY;
  BeginStyle := lsDiamond; // tipo da geometria do inicio da seta
  EndStyle := lsDiamond; // tipo da geometria do fim da seta
  BeginSize := 1; // tamanho da geometria do inicio da seta
  EndSize := 1; // tamanho da geometris do fim da seta
end;

// propriedades das setas que ligam os nós quando há dominancia pela HE
constructor TElasticEdge.Create3(AOwner: TSimpleGraph;
  SourceNode, TargetNode: TElasticNode);
begin
  inherited CreateNew(AOwner, SourceNode, [], TargetNode);
  Options := Options - [goSelectable];
  Pen.Color := clwebRed;
  BeginStyle := lsDiamond; // tipo da geometria do inicio da seta
  EndStyle := lsDiamond; // tipo da geometria do fim da seta
  BeginSize := 1; // tamanho da geometria do inicio da seta
  EndSize := 1; // tamanho da geometris do fim da seta
end;

procedure TfrmhasseD.QuebraPalavras(Texto: String);
var
  Lista: TStringList;
  i: Integer;
begin
  with TStringList.Create do
  begin
    Delimiter := ' ';
    DelimitedText := Texto;
    SetLength(TIWUserSession(WebApplication.Data).CountWords, 0);
    SetLength(TIWUserSession(WebApplication.Data).CountWords, Count);
  end;
end;

procedure TfrmhasseD.rgnHasseDiagramCreate(Sender: TObject);
begin
imghasse.Picture.LoadFromFile('C:\Users\CDCID 006 2017\Desktop\FITradeoff WEB Based - FU-TXMMO-WF1 (LAYOUT)\Images\Hasse\01' + '.jpg');
end;

function TfrmhasseD.SplitString(CountWords: array of String; Texto: string)
  : Integer;
var
  delimeter: string;
  idx, i: Integer;
  current_position: Integer;
  current_string: string;
begin
  delimeter := ' ';
  idx := 0;
  current_string := Texto;
  while true do
  begin
    current_position := Pos(delimeter, current_string);
    if current_position = 0 then // last item
    begin
      TIWUserSession(WebApplication.Data).CountWords[idx] := current_string;
      break;
    end;
    TIWUserSession(WebApplication.Data).CountWords[idx] := Copy
      (current_string, 1, current_position - 1);
    current_string := Copy(current_string, current_position + 1, Length
        (current_string) - current_position);
    inc(idx);
  end;

  for i := 0 to Length(TIWUserSession(WebApplication.Data).CountWords) - 1 do
    if Length(TIWUserSession(WebApplication.Data).CountWords[i]) > 7 then
    begin
      Result := 1;
      Exit;
    end
    else
      Result := 0;
end;

procedure TfrmhasseD.DomAlt2;
var
  i, j, o, k1, l1, ii, jj: Integer;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    k1 := 0;
    l1 := 0;
    SetLength(Node8, 0);
    SetLength(Node9, 0, 0);
    SetLength(Node9, Length(Node7), NumAlt);

    for o := aux_Nodei - 2 Downto 0 do
    begin
      // CASO SEJA A PRIMEIRA VEZ QUE RODA O PROCEDURE
      if o = aux_Nodei - 2 then
      begin
        // guarda no Node8 quais alt dominam a (alt)aux_Nodej da camada aux_Nodei
        for j := 0 to NumAlt - 1 do
        begin
          if Node5[CamAlt[aux_Nodei, aux_Nodej], j] > 0 then
          begin
            k1 := k1 + 1;
            SetLength(Node8, k1);
            Node8[k1 - 1] := Node5[CamAlt[aux_Nodei, aux_Nodej], j];
          end;
        end;

        // cria um vetor que será utilizado como auxílio na retirada dos valores do Node8
        SetLength(Node8Ant, Length(Node8));
        for i := 0 to Length(Node8) - 1 do
          Node8Ant[i] := Node8[i];

        // Cria a matriz que guarda quais alt dominam as alt presentes no Node8
        // Cada linha é uma alternativa
        for i := 0 to Length(Node7) - 1 do
        begin
          for j := 0 to NumAlt - 1 do
          begin
            if Node5[Node7[i] - 1, j] > 0 then
            begin
              Node9[i, l1] := Node5[Node7[i] - 1, j];
              l1 := l1 + 1;
            end // do if Node5[Node7[i], j] > 0
            else
              break;
          end; // do for j
          l1 := 0;
        end; // do for i

        // Tirar do Node8 as alt que não formarão ligações com a alt aux_Nodej da camada
        // aux_Nodei devido a transitividade
        /// ///////////////////////////////////////////////////////////////////////////
        for i := 0 to Length(Node8) - 1 do // roda as alt do Node8
        begin
          for j := 0 to Length(Node7) - 1 do
          // roda as linhas do Node9
          begin
            for k1 := 0 to NumAlt - 1 do
            // roda todas as colunas do Node9
            begin
              if ((Node8[i] = Node9[j, k1]) or (Node8[i] = Node7[j])) then
              begin
                Node8Ant[i] := 0;
                break;
              end;
              // do if Node8[i] = Node9[j, k]
            end; // do for k
          end; // do for j
        end; // do for i

        k1 := 0;
        l1 := -1;
        for i := 0 to Length(Node8Ant) - 1 do
        begin
          if Node8Ant[i] > 0 then
          begin
            k1 := k1 + 1;
            SetLength(Node8, k1);
            Node8[k1 - 1] := Node8Ant[i];
          end
          // do if Node8Ant[i] > 0
          else
          begin
            l1 := l1 + 1;
            if l1 = Length(Node8Ant) - 1 then
              SetLength(Node8, 0);
          end;
        end; // do for i

        /// /////////////////////////////////////////////////////////////////////////////
        l1 := 0;
        SetLength(Node7, 0);
        // //tentativa mudar cor Wesllen
        // if auxcontador =1 then
        // begin
        // for ii  := 0 to Numalt- 1 do
        // begin   //1
        // for jj := 0 to NumAlt - 1 do
        // if (Alternativas[jj] = auxavalholistic[1]) and (Alternativas[ii] = auxavalholistic[0]) then
        // begin//2
        // // criar a ligação entre as alt da camada i com a alt aux_Nodej da camada aux_Nodei
        //
        // for i := 0 to Length(Node8) - 1 do
        // begin//3
        // for j := 0 to Node3[o] - 1 do
        // begin //4
        // if Node8[i] - 1 = CamAlt[o, j] then
        // begin //5
        // if MatrizParaPar2[Node8[i] - 1, CamAlt[aux_Nodei, aux_Nodej]] = MatrizParapar2[i,j]
        // then
        // begin //6
        // TElasticEdge.Create1(SimpleGraph, Node[Node6[Node8[i] - 1]],
        // Node[Node6[CamAlt[aux_Nodei, aux_Nodej]]]);
        // l1 := l1 + 1;
        // SetLength(Node7, l1);
        // Node7[l1 - 1] := Node8[i];
        // break;
        // end;//6
        // // do if MatrizParaPar2[CamAlt[i - 1, k], CamAlt[i, j]] = 1
        // end; // do Node8[i] = CamAlt[i, j]    //5
        // end; // do for j //4
        // end; // do for i //3
        // end
        // else
        // begin //2
        //
        // //////////////////////////////////////////////////wesllen////////////////////////////
        // // criar a ligação entre as alt da camada i com a alt aux_Nodej da camada aux_Nodei
        //
        // // for i := 0 to Length(Node8) - 1 do
        // // begin
        // // for j := 0 to Node3[o] - 1 do
        // // begin
        // //   if Node8[i] - 1 = CamAlt[o, j] then
        // // begin
        // if MatrizParaPar2[Node8[i] - 1, CamAlt[aux_Nodei, aux_Nodej]] = 1
        // then
        // begin
        // TElasticEdge.Create(SimpleGraph, Node[Node6[Node8[i] - 1]],
        // Node[Node6[CamAlt[aux_Nodei, aux_Nodej]]]);
        // l1 := l1 + 1;
        // SetLength(Node7, l1);
        // Node7[l1 - 1] := Node8[i];
        // break;
        // end;
        // // do if MatrizParaPar2[CamAlt[i - 1, k], CamAlt[i, j]] = 1
        // end; // do Node8[i] = CamAlt[i, j]
        // end; // do for j
        // //  end; // do for i
        // //    end;
        // //   end;
        // end
        // else
        // criar a ligação entre as alt da camada i com a alt aux_Nodej da camada aux_Nodei

        for i := 0 to Length(Node8) - 1 do
        begin
          for j := 0 to Node3[o] - 1 do
          begin
            if Node8[i] - 1 = CamAlt[o, j] then
            begin
              if MatrizParaPar2[Node8[i] - 1, CamAlt[aux_Nodei, aux_Nodej]]
                = 1 then
              begin
                TElasticEdge.Create(SimpleGraph, Node[Node6[Node8[i] - 1]],
                  Node[Node6[CamAlt[aux_Nodei, aux_Nodej]]]);
                l1 := l1 + 1;
                SetLength(Node7, l1);
                Node7[l1 - 1] := Node8[i];
                break;
              end;
              // do if MatrizParaPar2[CamAlt[i - 1, k], CamAlt[i, j]] = 1
            end; // do Node8[i] = CamAlt[i, j]
          end; // do for j
        end; // do for i

        if Length(Node7) = 0 then
          break;
      end // do if o = aux_Nodei - 2

      // CASO O PROCEDURE JÁ TENHA RODADO A PARTE ACIMA
      else
      begin

        // cria um vetor que será utilizado como auxílio na retirada dos valores do Node8
        SetLength(Node8Ant, Length(Node8));
        for i := 0 to Length(Node8) - 1 do
          Node8Ant[i] := Node8[i];

        // Cria a matriz que guarda quais alt dominam as alt presentes no Node8
        // Cada linha é uma alternativa
        SetLength(Node9, 0, 0);
        SetLength(Node9, Length(Node7), NumAlt);
        l1 := 0;
        for i := 0 to Length(Node7) - 1 do
        begin
          for j := 0 to NumAlt - 1 do
          begin
            if Node5[Node7[i] - 1, j] > 0 then
            begin
              Node9[i, l1] := Node5[Node7[i] - 1, j];
              l1 := l1 + 1;
            end // do if Node5[Node7[i], j] > 0
            else
              break;
          end; // do for j
          l1 := 0;
        end; // do for i

        // Tirar do Node8 as alt que não formarão ligações com a alt aux_Nodej da camada
        // aux_Nodei devido a transitividade
        /// ///////////////////////////////////////////////////////////////////////////
        for i := 0 to Length(Node8) - 1 do // roda as alt do Node8
        begin
          for j := 0 to Length(Node7) - 1 do
          // roda as linhas do Node9
          begin
            for k1 := 0 to NumAlt - 1 do // roda as colunas do Node9
            begin
              if ((Node8[i] = Node9[j, k1]) or (Node8[i] = Node7[j])) then
              begin
                Node8Ant[i] := 0;
                break;
              end;
              // do if Node8[i] = Node9[j, k]
            end; // do for k
          end; // do for j
        end; // do for i

        k1 := 0;
        l1 := -1;
        for i := 0 to Length(Node8Ant) - 1 do
        begin
          if Node8Ant[i] > 0 then
          begin
            k1 := k1 + 1;
            SetLength(Node8, k1);
            Node8[k1 - 1] := Node8Ant[i];
          end
          // do if Node8Ant[i] > 0
          else
          begin
            l1 := l1 + 1;
            if l1 = Length(Node8Ant) - 1 then
              SetLength(Node8, 0);
          end;
        end; // do for i

        /// /////////////////////////////////////////////////////////////////////////////

        // criar a ligação entre as alt da camada i com a alt aux_Nodej da camada aux_Nodei
        l1 := 0;
        SetLength(Node7, 0);
        for i := 0 to Length(Node8) - 1 do
        begin
          for j := 0 to Node3[o] - 1 do
          begin
            if Node8[i] - 1 = CamAlt[o, j] then
            begin
              if MatrizParaPar2[Node8[i] - 1, CamAlt[aux_Nodei, aux_Nodej]]
                = 1 then
              begin
                TElasticEdge.Create(SimpleGraph, Node[Node6[Node8[i] - 1]],
                  Node[Node6[CamAlt[aux_Nodei, aux_Nodej]]]);
                l1 := l1 + 1;
                SetLength(Node7, l1);
                Node7[l1 - 1] := Node8[i];
                break;
              end;
              // do if MatrizParaPar2[CamAlt[i - 1, k], CamAlt[i, j]] = 1
            end; // do Node8[i] = CamAlt[i, j]
          end; // do for j
        end; // do for i
      end; // do else do if o = aux_Nodei - 2

      if Length(Node7) = 0 then
        break;

    end; // do for o

  end; // with
end;

procedure TfrmhasseD.AMD;
var
  i, j, k, l, m, n, cont, g, o: Integer; // variáveis auxiliares
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    if contador2 = 0 then // criação da primeira camada
    begin
      SetLength(Node2Ant, NumAlt);

      for i := 0 to Length(Node2) - 1 do
        Node2Ant[i] := Node2[i];

      SetLength(Node1, 0);
      cont := 0;
      k := 0;
      l := 0;

      for j := 0 to NumAlt - 1 do
      begin
        for i := 0 to NumAlt - 1 do
        begin
          if (MatrizParaPar2[i, j] = 0) or (MatrizParaPar2[i, j] = 2) then
          begin
            cont := cont + 1;
            if cont = NumAlt then
            begin
              k := k + 1;
              SetLength(Node1, k);
              Node1[l] := j;
              l := l + 1;
            end;
          end
          else
            break;
        end;
        cont := 0;
      end;
      contador := contador + 1;
      contador2 := contador2 + 1;
    end // do contador2 = 0
    else // caso já tenha feita a primeira camada
    begin
      SetLength(Node2, Length(Node2Ant) - Length(Node1));

      m := 1;
      k := 0;
      l := 0;

      // guarda as alternativas que faltam
      for i := 0 to Length(Node2Ant) - 1 do
      begin
        if Node2Ant[i] = Node1[l] then
        begin
          k := k + 1;
          if l = Length(Node1) - 1 then
          begin
            l := (Length(Node1) - 1);
          end
          else
            l := l + 1;
          Continue;
        end
        else
        begin
          SetLength(Node2, m);
          Node2[m - 1] := Node2Ant[k];
          m := m + 1;
          k := k + 1;
        end;
      end;

      SetLength(Node2Ant, Length(Node2));
      for i := 0 to Length(Node2) - 1 do
        Node2Ant[i] := Node2[i];

      SetLength(Node1, 0);
      cont := 0;
      k := 0;
      l := 0;
      m := 0;
      n := 0;

      if Length(Node2) = 1 then
      begin
        k := k + 1;
        SetLength(Node1, k);
        Node1[l] := Node2[l];
      end
      else
      begin
        for j := 0 to Length(Node2) - 1 do
        begin
          g := Node2[m];
          for i := 0 to Length(Node2) - 1 do
          begin
            o := Node2[n];
            if (MatrizParaPar2[o, g] = 0) or (MatrizParaPar2[o, g] = 2) then
            begin
              cont := cont + 1;
              n := n + 1;
              if cont = Length(Node2) then
              begin
                k := k + 1;
                SetLength(Node1, k);
                Node1[l] := g;
                l := l + 1;
              end; // do if cont
            end
            else // do if MatrizParaPar2[i, j]
              break;
          end; // do for i
          cont := 0;
          n := 0;
          m := m + 1; // do if = Node2

        end; // do for j
      end; // do if length(Node2) = 1
      contador := contador + 1;
      contador2 := contador2 + 1;
    end;
  end; // WITH

end;

procedure TfrmhasseD.IWButton1Click(Sender: TObject);
var

  Node: Array of TElasticNode; // nós do grafo
  i, j, k, a, b, c, d, l: Integer;
  Result: Set of TGraphObjectOption;
  Pt: array [0 .. 1] of TPoint;
  cont: Integer; // índice da camada
begin

  with TIWUserSession(WebApplication.Data) do
  begin
    SetLength(MatrizParaPar2, NumAlt, NumAlt);
    SetLength(CamAlt, NumAlt, NumAlt);
    SetLength(Node2, NumAlt);
    SetLength(Node8, 0);

    for i := 0 to NumAlt - 1 do
    begin
      for j := 0 to NumAlt - 1 do
        MatrizParaPar2[i, j] := ArmazenaParaPar[Ciclo - 1, i, j];
    end;

    for i := 0 to NumAlt - 1 do
      Node2[i] := i;

    // ** Verifica a relação de dominância entre as alternativas **//
    SetLength(Node5, NumAlt, NumAlt);
    k := 0;
    for j := 0 to NumAlt - 1 do
    begin
      for i := 0 to NumAlt - 1 do
      begin
        if MatrizParaPar2[i, j] = 1 then
        begin
          Node5[j, k] := i + 1;
          k := k + 1;
        end; // do if
      end; // do for i
      k := 0;
    end; // do for j
    /// ///////////////////////////////////////////////////////////////////////////

    // ** Verifica a quais camadas cada Alternativa pertence **//
    SetLength(CamAlt, 0, 0); // zerando matriz
    SetLength(CamAlt, NumAlt, NumAlt);
    SetLength(Node3, 0);

    QntCam := 0;
    k := 0;

    for j := 0 to NumAlt - 1 do
    begin
      if (Length(Node2) - Length(Node1)) = 0 then
      begin
        Continue;
      end
      else
      begin
        AMD;
        QntCam := QntCam + 1;
        for i := 0 to Length(Node1) - 1 do
        begin
          SetLength(Node3, QntCam);
          CamAlt[contador - 1, i] := Node1[i];
          k := k + 1;
        end;
        Node3[j] := k;
        k := 0;
      end;
    end;
    /// /////////////////////////////////////////////////////////////////////////////

    // ** Criando a posição dos nós **//
    SetLength(Node, NumAlt);

    k := 0;
    c := trunc(400 / (QntCam + 1));
    b := -400; // posição inicial y

    for i := 0 to QntCam - 1 do
    begin
      if Node3[i] = 1 then
      begin
        b := b + 100;
        Node[k] := TElasticNode.Create(SimpleGraph, 0, b);
      end // do if Node3[i] = 1
      else
      begin
        if Odd(Node3[i]) then // caso a quantidade de alt na camada for IMPAR
        begin
          a := -(trunc((Node3[i] - 1) / 2)) * 200;
          b := b + 100;
          for j := 0 to Node3[i] - 1 do
          begin
            Node[j + k] := TElasticNode.Create(SimpleGraph, a, b);
            a := a + 200;
          end;
        end
        else // caso seja PAR
        begin
          if Node3[i] = 2 then
          begin
            a := -100;
            b := b + 100;
            for j := 0 to Node3[i] - 1 do
            begin
              Node[j + k] := TElasticNode.Create(SimpleGraph, a, b);
              a := a + 200;
            end; // do for j
          end // do if Node3[i] = 2
          else
          begin
            a := -(((trunc(Node3[i] / 2)) - 1) * 200 + 100);
            b := b + 100;
            for j := 0 to Node3[i] - 1 do
            begin
              Node[j + k] := TElasticNode.Create(SimpleGraph, a, b);
              a := a + 200;
            end;
            // do for j
          end; // do else do if Node3[i] = 2
        end; // do caso seja PAR
      end; // else do if Node3[i] = 1
      k := k + Node3[i];
    end; // do for i
    /// ///////////////////////////////////////////////////////////////////////////

    // ** Nome dos Nodos **//
    Result := [];
    Include(Result, goShowCaption);

    SetLength(AltOrdemGrafo, NumAlt);
    k := 0;

    for i := 0 to QntCam - 1 do
    begin
      for j := 0 to Node3[i] - 1 do
      begin
        AltOrdemGrafo[j + k] := CamAlt[i, j];
      end;
      k := k + j;
    end;

    for i := 0 to NumAlt - 1 do
    begin
      with Node[i] do
      begin
        Pen.Color := clwebblack;
        Font.Style := [fsBold];
        text := NomeAlt[AltOrdemGrafo[i]];
        QuebraPalavras(NomeAlt[AltOrdemGrafo[i]]);
        if ((Length(NomeAlt[AltOrdemGrafo[i]]) > 23) or
            (SplitString(TIWUserSession(WebApplication.Data).CountWords,
              NomeAlt[AltOrdemGrafo[i]]) = 1)) then
          Font.Size := 7;
        // else                    //inc 09.05 inspire
        // Font.Size := 20;      //inc 09.05 inspire
      end;
    end;
    /// /////////////////////////////////////////////////////////////////////////////

    // ** NumNode **//
    SetLength(Node6, NumAlt);
    SetLength(Node7, 0);
    l := 0;

    for i := 0 to NumAlt - 1 do
    begin
      for j := 0 to NumAlt - 1 do
      begin
        if AltOrdemGrafo[j] = i then
        begin
          Node6[i] := j;
          break;
        end;
      end;
    end;

    for i := QntCam - 1 Downto 1 do
    begin
      for j := 0 to Node3[i] - 1 do
      // qnt de elementos da camada (i-1)
      begin
        for k := 0 to Node3[i - 1] - 1 do
        begin
          if MatrizParaPar2[CamAlt[i - 1, k], CamAlt[i, j]] = 1 then
          // begin
          // if auxContador = 1 then
          begin
            if matrizholistica[CamAlt[i - 1, k], CamAlt[i, j]] = 1 then
            begin
              TElasticEdge.Create1(SimpleGraph, Node[Node6[CamAlt[i - 1, k]]],
                Node[Node6[CamAlt[i, j]]]);
            end
            else
            begin
              TElasticEdge.Create(SimpleGraph, Node[Node6[CamAlt[i - 1, k]]],
                Node[Node6[CamAlt[i, j]]]);
            end;

            // end
            // else
            // begin
            // TElasticEdge.Create(SimpleGraph, Node[Node6[CamAlt[i - 1, k]]],
            // Node[Node6[CamAlt[i, j]]]);
            // end;
            l := l + 1;
            SetLength(Node7, l);
            Node7[l - 1] := CamAlt[i - 1, k] + 1;
          end; // do if
        end; // do for k
        aux_Nodei := i;
        aux_Nodej := j;
        DomAlt2(SimpleGraph, Node);
        l := 0;
      end; // do for j
    end; // do for i
    /// /////////////////////////////////////////////////////////////////////////////

    // ** NumNode2 **//
    for i := 0 to NumAlt - 1 do
    begin
      for j := 0 to NumAlt - 1 do
      begin
        if auxcontador = 1 then
        begin

        end;
        if MatrizParaPar2[i, j] = 2 then
        begin
          TElasticEdge.Create2(SimpleGraph, Node[Node6[i]], Node[Node6[j]]);
        end; // do if
      end; // do for j
    end; // do for i
    /// ///////////////////////////////////////////////////////////////////////////

    // ** Linha Tracejada **//
    cont := Rlevels - 1;
    aux_i := 0;
    for i := QntCam - 1 Downto 1 do // camada do elemento analisado para ser dominado
    begin
      // i := QntCam - 1;
      aux_i := 0;
      for j := 0 to Node3[i] - 1 do
      // elemento analisado para ser dominado
      begin
        for k := i - 1 Downto 0 do // camada do elemento que domina
        begin
          for l := 0 to Node3[k] - 1 do // elemento que domina
          begin
            if MatrizParaPar2[CamAlt[k, l], CamAlt[i, j]] = 1 then
            begin
              if ((j = Node3[i] - 1) and (k = 0) and (l = Node3[0] - 1)) then
              begin
                a := 100 * i - 10;
                Pt[0] := Point(100, a);
                Pt[1] := Point(1100, a);
                with SimpleGraph.InsertLink(Pt, nil) do
                begin
                  Pen.Style := psDot;
                  EndStyle := lsNone;
                end;
                // criar a legenda dos groups
                Pt[0] := Point(980, a + 95);
                Pt[1] := Point(1050, a + 95);
                with SimpleGraph.InsertLink(Pt, nil) do
                begin
                  Pen.Style := psSolid;
                  EndStyle := lsNone;
                  Font.Color := clwebblack;
                  Font.Style := [fsBold];
                  text := 'Position' + FloatTostr(cont);
                  cont := cont - 1;
                end;
                aux_i := 1;
              end;
            end
            else
            begin
              aux_i := 1;
              break;
            end;
            if aux_i = 1 then
              break;
          end; // do for l
          if aux_i = 1 then
            break;
        end; // do for k
        if aux_i = 1 then
          break;
      end; // do for j
    end; // do for i

    Pt[0] := Point(980, a - 5);
    Pt[1] := Point(1050, a - 5);
    with SimpleGraph.InsertLink(Pt, nil) do
    begin
      Pen.Style := psSolid;
      EndStyle := lsNone;
      Font.Color := clwebblack;
      Font.Style := [fsBold];
      text := 'Position 1';
    end;

    SimpleGraph.LockLinks := true;

  end;

end;

procedure TfrmhasseD.IWButton2Click(Sender: TObject);
begin
  with TIWUserSession(WebApplication.Data) do
  begin
  SimpleGraph.Create(rgnHasseDiagram);
  SimpleGraph.Height := 609;
  SimpleGraph.Width := 1200;
  SimpleGraph.Top := 192;
  SimpleGraph.Left := 16;


  end;
end;

initialization

TfrmhasseD.SetAsMainForm;

end.
