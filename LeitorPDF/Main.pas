{
Data: 23/08/2026
Leitor de arquivo PDF sem permissão de copiar e imprimir.
Desenvolvido por: Leandro Silva 0,1
email: leandrox364@gmail.com
}
unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls,
  Vcl.Imaging.pngimage, Vcl.Buttons, Vcl.Imaging.jpeg;

type
  TFrmMain = class(TForm)
    PanelBotoes: TPanel;
    OpenDialog1: TOpenDialog;
    BtnCarregarPDF: TBitBtn;
    BtnPagAnterior: TBitBtn;
    BtnProximaPag: TBitBtn;
    ImageLogo: TImage;
    ScrollBox1: TScrollBox;
    imgPagina: TImage;
    BtnZoomOut: TBitBtn;
    BtnZoomIN: TBitBtn;
    pnlContainer: TPanel;
    LabelNomeArquivo: TLabel;
    procedure BtnCarregarPDFClick(Sender: TObject);
    procedure BtnPagAnteriorClick(Sender: TObject);
    procedure BtnProximaPagClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BtnZoomINClick(Sender: TObject);
    procedure BtnZoomOutClick(Sender: TObject);
    procedure FormMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure ImageLogoClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
  private
    { Private declarations }
    FCurrentPage: Integer;
    FPdfPath: string;
    FTempDir: string;
    FZoomPercent: Integer; // Guardará valores como 100, 120, 80, etc.

    function  ExecutarOculto(Cmd: string): Boolean;
    procedure RenderizarPagina;
    procedure AplicarZoom(Modificador: Integer);
    procedure CentralizarImagem;

  public
    { Public declarations }
     FZoomFactor: Double;
  end;
const
   VERSAO = '1.0';

var
  FrmMain: TFrmMain;

implementation

{$R *.dfm}

function TFrmMain.ExecutarOculto(Cmd: string): Boolean;
var
  SUInfo: TStartupInfo;
  PInfo: TProcessInformation;
begin
  FillChar(SUInfo, SizeOf(SUInfo), 0);
  SUInfo.cb := SizeOf(SUInfo);
  SUInfo.dwFlags := STARTF_USESHOWWINDOW;
  SUInfo.wShowWindow := SW_HIDE; // Mantém a tela preta do prompt oculta

  Result := CreateProcess(nil, PChar(Cmd), nil, nil, False,
                          CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS,
                          nil, nil, SUInfo, PInfo);
  if Result then
  begin
    // Aguarda o término da renderização do pdftopng antes de seguir no Delphi
    WaitForSingleObject(PInfo.hProcess, INFINITE);
    CloseHandle(PInfo.hProcess);
    CloseHandle(PInfo.hThread);
  end;
end;

procedure TFrmMain.FormCreate(Sender: TObject);
begin
  FZoomFactor := 1.0; // 100% do tamanho original

  FPdfPath := ParamStr(1);

  if Pos('doc' , FPdfPath ) > 0 then
  begin
    ShowMessage('Não é possível abrir arquivo do Word.');
    Application.Terminate;
    Exit;
  end;
//  ShowMessage('Arquivo: '+ FPdfPath);
  if FPdfPath  <> '' then
  begin
    BtnCarregarPDF.Visible :=False;
    FCurrentPage := 1;
    RenderizarPagina;
    FZoomPercent := 100;

    LabelNomeArquivo.Caption:= ' ' + ExtractFileName(FPdfPath);
    Caption := ExtractFileName(FPdfPath);
  end;
end;

procedure TFrmMain.FormMouseWheel(Sender: TObject; Shift: TShiftState;
  WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
var
  LPosAntiga: Integer;
begin
  LPosAntiga := ScrollBox1.VertScrollBar.Position;

  if WheelDelta > 0 then
  begin
    ScrollBox1.VertScrollBar.Position := ScrollBox1.VertScrollBar.Position - 30;

    // Se tentou subir e a barra já estava no topo (posição 0), vai para a página anterior
    if (LPosAntiga = 0) and (ScrollBox1.VertScrollBar.Position = 0) then
    begin
      if BtnPagAnterior.Enabled then
      begin
         BtnPagAnterior.Click;
      // Joga o scroll para o fim da página anterior que acabou de carregar
           ScrollBox1.VertScrollBar.Position := ScrollBox1.VertScrollBar.Range;
      end;
    end;
  end
  else
  begin
    ScrollBox1.VertScrollBar.Position := ScrollBox1.VertScrollBar.Position + 30;

    // Se tentou descer e a posição não mudou mais, significa que chegou no fim da folha
    if LPosAntiga = ScrollBox1.VertScrollBar.Position then
    begin
      if BtnProximaPag.Enabled then
      begin
         BtnProximaPag.Click;
      // Joga o scroll de volta para o topo da nova página carregada
         ScrollBox1.VertScrollBar.Position := 0;
      end;
    end;
  end;

  Handled := True;
end;


procedure TFrmMain.FormResize(Sender: TObject);
begin
  if Width <= 1024 then
     Width := 1024;

  CentralizarImagem;
end;

procedure TFrmMain.ImageLogoClick(Sender: TObject);
begin
  ShowMessage('Leitor de Arquivo PDF' + #13 +
              'Versão: '+ VERSAO      + #13 +
              'Desenvolvido: Leandro SSilva {0,1}.');
end;

procedure TFrmMain.RenderizarPagina;
var
  Cmd: string;
  ImgBase: string;
  ImgGerada: string;
  ExePath: string;
begin
  if FPdfPath = '' then Exit;

  FTempDir := ExtractFilePath(Application.ExeName);


  ForceDirectories(FTempDir);

  ExePath :=   ExtractFilePath(Application.ExeName)+'pdftopng.exe';
  ImgBase := FTempDir + 'pagina_atual';

  if not FileExists(ExePath) then
     raise Exception.Create('Não foi encontrado o gerado de arquivo. pdftopng.exe');


  // O pdftopng obrigatoriamente coloca traço e 6 dígitos de numeração
  ImgGerada := Format('%s-%0.6d.png', [ImgBase, FCurrentPage]);

  // Limpa o cache antigo antes de chamar o comando
  if FileExists(ImgGerada) then
     DeleteFile(ImgGerada);

// linha de comando 'D:\Projetos\LeitorPDF\pdftopng.exe -f 1 -l 1
//                  "D:\Projetos\LeitorPDF\Word.pdf"
//                  "D:\Projetos\LeitorPDF\pagina_atual"'

  // Monta a linha de execução com os parâmetros
  Cmd := Format('%s -f %d -l %d "%s" "%s"',
                [ExePath, FCurrentPage, FCurrentPage, FPdfPath, ImgBase]);

  ExecutarOculto(Cmd);

  // Verifica se o arquivo físico real foi gravado na pasta
  if FileExists(ImgGerada) then
  begin
    try
       imgPagina.Picture.LoadFromFile(ImgGerada);

     // Atualiza os gráficos da tela
      imgPagina.Refresh;
      imgPagina.AutoSize := True; // Reseta para ler o tamanho nativo do novo arquivo PNG
      AplicarZoom(0); // Aplica o percentual atual (FZoomPercent) na nova página carregada
      CentralizarImagem;

    finally
      DeleteFile(ImgGerada);
    end;
  end
  else
  begin
    if FCurrentPage > 1 then
    begin
      BtnProximaPag.Enabled:=False;
      BtnPagAnterior.Enabled:=True;

      Dec(FCurrentPage);
      ShowMessage('Fim do documento ou erro ao gerar a página.');

    end;
  end;
end;
procedure TFrmMain.BtnCarregarPDFClick(Sender: TObject);
begin
  OpenDialog1.Filter := 'Arquivos PDF|*.pdf';
  OpenDialog1.InitialDir := ExtractFilePath(Application.ExeName);
  if OpenDialog1.Execute then
  begin
    FPdfPath := OpenDialog1.FileName;
    FCurrentPage := 1;
    RenderizarPagina;
    FZoomPercent := 100;
    LabelNomeArquivo.Caption:= ' ' + ExtractFileName(FPdfPath);
  end;
end;

procedure TFrmMain.BtnPagAnteriorClick(Sender: TObject);
begin
  if FCurrentPage > 1 then
  begin
    Dec(FCurrentPage);
    RenderizarPagina;
    BtnProximaPag.Enabled :=True;
    // Joga o scroll para o fim da página anterior que acabou de carregar
    ScrollBox1.VertScrollBar.Position := ScrollBox1.VertScrollBar.Range;

  end
  else
  begin
    BtnPagAnterior.Enabled :=false;
    BtnProximaPag.Enabled :=True;
  end;

end;

procedure TFrmMain.BtnProximaPagClick(Sender: TObject);
begin
  Inc(FCurrentPage);
  RenderizarPagina;
  BtnPagAnterior.Enabled :=True;
  // Joga o scroll para o inicio da página
  ScrollBox1.VertScrollBar.Position := 0;

end;

procedure TFrmMain.BtnZoomINClick(Sender: TObject);
begin
  AplicarZoom(-10); // Diminui 10%
end;

procedure TFrmMain.BtnZoomOutClick(Sender: TObject);
begin
  AplicarZoom(10); // Aumenta 10%
end;

procedure TFrmMain.AplicarZoom(Modificador: Integer);
var
  NovoPercent: Integer;
  LarguraOriginal, AlturaOriginal: Integer;
begin
  NovoPercent := FZoomPercent + Modificador;

  // Limita o zoom entre 30% (mínimo) e 300% (máximo) para evitar distorções absurdos
  if (NovoPercent < 80) or (NovoPercent > 300) then Exit;

  FZoomPercent := NovoPercent;

  // Desativa temporariamente o AutoSize para podermos forçar o novo tamanho
  imgPagina.AutoSize := False;
  imgPagina.Stretch := True;

  // Primeiro descobrimos o tamanho real da imagem recarregando os metadados dela
  LarguraOriginal := imgPagina.Picture.Width;
  AlturaOriginal  := imgPagina.Picture.Height;

  // Calcula matematicamente o novo tamanho proporcional baseado no percentual
  imgPagina.Width  := MulDiv(LarguraOriginal, FZoomPercent, 100);
  imgPagina.Height := MulDiv(AlturaOriginal, FZoomPercent, 100);

  CentralizarImagem;

end;


procedure TFrmMain.CentralizarImagem;
var
  NovaLargura, NovaAltura: Integer;
begin
  // 1. Define o tamanho do painel baseado no maior valor (tela ou tamanho da imagem)
  if imgPagina.Width < ScrollBox1.ClientWidth then
    NovaLargura := ScrollBox1.ClientWidth
  else
    NovaLargura := imgPagina.Width;

  if imgPagina.Height < ScrollBox1.ClientHeight then
    NovaAltura := ScrollBox1.ClientHeight
  else
    NovaAltura := imgPagina.Height;

  // Aplica o tamanho calculado ao contêiner intermediário
  pnlContainer.Width := NovaLargura;
  pnlContainer.Height := NovaAltura;

  // 2. Centraliza a imagem dentro do contêiner se ela for menor que a área visível
  if imgPagina.Width < ScrollBox1.ClientWidth then
    imgPagina.Left := (ScrollBox1.ClientWidth - imgPagina.Width) div 2
  else
    imgPagina.Left := 0;

  if imgPagina.Height < ScrollBox1.ClientHeight then
    imgPagina.Top := (ScrollBox1.ClientHeight - imgPagina.Height) div 2
  else
    imgPagina.Top := 0;
end;
end.
