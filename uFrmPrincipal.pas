unit uFrmPrincipal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, SqlExpr, IdBaseComponent, IdComponent, System.UITypes,
  IdTCPConnection, IdTCPClient, IdFTP, ExtCtrls, ComCtrls, StdCtrls, Vcl.FileCtrl,
  Buttons, Gauges, IniFiles, WinInet, WideStrings, IdExplicitTLSClientServerBase,
  FMTBcd, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL, IdSSLOpenSSL,
  IdIntercept, IdLogBase, IdLogFile, IdServerIOHandler, Data.DBXMsSQL;

const 
  gsArqRemoto = 'Atualizacoes.amk'; 
  gsArqLocal = 'Versoes.amk';

type
  TFrmPrincipal = class(TForm)
    BvLinha: TBevel;
    Panel2: TPanel;
    lblListandoAtalizacoes: TLabel;
    lblTotalBytes: TLabel;
    pnlNomeModulo: TPanel;
    lblStatus: TLabel;
    lblContador: TLabel;
    lblNomeModulo: TLabel;
    GgProgressao: TGauge;
    btnAtualizar: TBitBtn;
    btnFechar: TBitBtn;
    lstModulos: TListView;
    TmVerificacao: TTimer;
    idFtpAtualiza: TIdFTP;
    MmScripts: TMemo;
    LbArquivosRemotos: TListBox;
    SQLConnConexao: TSQLConnection;
    IdSSLIOHandlerSocketOpenSSL1: TIdSSLIOHandlerSocketOpenSSL;
    procedure btnFecharClick(Sender: TObject);
    procedure TmVerificacaoTimer(Sender: TObject);
    procedure btnAtualizarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure idFtpAtualizaWorkBegin(ASender: TObject; AWorkMode: TWorkMode;
      AWorkCountMax: Int64);
    procedure idFtpAtualizaWork(ASender: TObject; AWorkMode: TWorkMode;
      AWorkCount: Int64);
  private
    procedure ApagarTemporarios;
    procedure LimparPastaTemp;
    procedure AtualizarINIVersoes(ASecao, ANomeVersao, AVersao: string);
  public
    flagAtivado         : boolean;
    gsPastaRemota       : string;
    Tamanho_Arquivo     : LongWord;
    STime               : TDateTime;
    Tempo_Medio         : Double;
    Bytes_Transf        : Double;
    ModuloEmAtualizacao : string;
    gsPastaLocal        : string;

    sArqLocal           : string;
    sArqRemota          : string;
    Modulo              : string;
    Titulo              : string;
    ModulosAbertos      : TStringList;
    IniLocal            : TIniFile;
    IniRemota           : TIniFile;
    VersaoLocal         : string;
    VersaoRemota        : string;
    kbTotalDown         : LongWord;
    procedure CriaArquivoVersoes;
    procedure PreencherModulosAbertos;
    function  Conectar: Boolean;
    procedure LimpaLabels;
    procedure IncluiDownloads(AModulo, AVersaoAtual, AVersaoRemota: string;
      ATamanho: LongWord);
    procedure PosAtualizacao;

    procedure conexaoBancoDados;
    procedure executarScripts;
    procedure GetValorArquivoIni( Secao, NomeVariavel : string;
      var Resultado : string );
    function executaSQL( sql : string ) : boolean;
  end;

var
  FrmPrincipal: TFrmPrincipal;

implementation

{$R *.dfm}

procedure TFrmPrincipal.ApagarTemporarios;
var
  lpEntryInfo : PInternetCacheEntryInfo;
  hCacheDir   : LongWord;
  dwEntrySize : LongWord;
begin
  dwEntrySize := 0;
  FindFirstUrlCacheEntry(nil, TInternetCacheEntryInfo(nil^), dwEntrySize);
  GetMem(lpEntryInfo, dwEntrySize);
  if dwEntrySize > 0 then lpEntryInfo^.dwStructSize := dwEntrySize;
  hCacheDir := FindFirstUrlCacheEntry(nil, lpEntryInfo^, dwEntrySize);
  if hCacheDir <> 0 then
  begin
    repeat
      DeleteUrlCacheEntry(lpEntryInfo^.lpszSourceUrlName);
      FreeMem(lpEntryInfo, dwEntrySize);
      dwEntrySize := 0;
      FindNextUrlCacheEntry(hCacheDir, TInternetCacheEntryInfo(nil^), dwEntrySize);
      GetMem(lpEntryInfo, dwEntrySize);
      if dwEntrySize > 0 then lpEntryInfo^.dwStructSize := dwEntrySize;
    until not FindNextUrlCacheEntry(hCacheDir, lpEntryInfo^, dwEntrySize);
  end;
  FreeMem(lpEntryInfo, dwEntrySize);
  FindCloseUrlCache(hCacheDir);
end;

procedure TFrmPrincipal.AtualizarINIVersoes(ASecao, ANomeVersao,
  AVersao: string);
var
  Ini : TIniFile;
begin
  Ini := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'Versoes.amk');
  Ini.WriteString(ASecao, ANomeVersao, AVersao);
  Ini.Free;
end;

procedure TFrmPrincipal.btnFecharClick(Sender: TObject);
begin
  Close;
end;

function TFrmPrincipal.Conectar: Boolean;
begin
  try
    with idFtpAtualiza do
    begin
      gsPastaRemota := '/softwares/cartorio/';
      gsPastaLocal := ExtractFilePath(Application.ExeName);
      sArqRemota := gsPastaRemota + '/' + gsArqRemoto;
      sArqLocal := gsPastaLocal + gsArqRemoto;
      Host := '191.252.200.2';
      Username := 'user_ftp';
      Password := 'Thi@go28';
      Passive := True;

      if not Connected then
        Connect;

      {Desativa temporariamente os eventos do idFTP pra não mostrar o download do arquivo texto}
      OnWorkBegin := nil;
      OnWork := nil;
      flagAtivado := false;

      try
        ChangeDir(gsPastaRemota);
      except
        TmVerificacao.Enabled := False;
        Halt;
      end;

      List(nil);
      result := True;
    end;
  except
    result := False;
  end;
end;

procedure TFrmPrincipal.conexaoBancoDados;
var
   Resultado : string;
begin
  GetValorArquivoIni( 'BANCODADOS' , 'DATABASE' , Resultado );

  SQLConnConexao.ConnectionName := 'Sistema';
  SQLConnConexao.DriverName := 'MSSQL';
  SQLConnConexao.GetDriverFunc := 'getSQLDriverMSSQL';
  SQLConnConexao.LibraryName := 'dbexpmss.dll';
  SQLConnConexao.VendorLib := 'sqlncli10.dll';

  SQLConnConexao.Params.Values[ 'HostName'] := resultado;
  SQLConnConexao.Params.Values[ 'DataBase' ] := 'CART';
  SQLConnConexao.Params.Values[ 'User_Name' ] := 'sa';
  SQLConnConexao.Params.Values[ 'Password' ] := 'M@rcus56';
  SQLConnConexao.Params.Values[ 'OS Authentication' ] := 'False';
  SQLConnConexao.Params.Values[ 'Mars_Connection' ] := 'True';

  SQLConnConexao.Connected := true;
end;

procedure TFrmPrincipal.CriaArquivoVersoes;
var
  StringList : TStringList;
begin
  StringList := TStringList.Create;
  with StringList do
    begin
      Add('[Atualizacoes]');
      Add('Modulo1=Administrativo.exe');
      Add('Versao1=vF1.1.01');
      Add('Titulo1=Gerenciamento do sistema');
      Add('');
      Add('Modulo2=Balcao.exe');
      Add('Versao2=vF1.1.01');
      Add('Titulo2=Serviços do balcão');
      Add('');
      Add('Modulo3=RegistroTitulosDocumentos.exe');
      Add('Versao3=vF1.1.01');
      Add('Titulo3=Registro de títulos e documentos');
      Add('');
      Add('Modulo4=Tabelionato.exe');
      Add('Versao4=vF1.1.01');
      Add('Titulo4=Escritura');
      Add('');
      Add('Modulo5=ProtestoTitulos.exe');
      Add('Versao5=vF1.1.01');
      Add('Titulo5=Protesto de títulos e documentos');
      Add('');
      Add('Modulo6=RegistroGeralImoveis.exe');
      Add('Versao6=vF1.1.01');
      Add('Titulo6=Registro geral de imóveis');
      Add('');
      Add('Modulo7=RegistroPessoaJuridica.exe');
      Add('Versao7=vF1.1.01');
      Add('Titulo7=Registro de pessoa jurídica');
      Add('');
      Add('Modulo8=RegistroPessoasNaturais.exe');
      Add('Versao8=vF1.1.01');
      Add('Titulo8=Registro de pessoas naturais');
      Add('');
      Add('Modulo9=InterdicaoTutelas.exe');
      Add('Versao9=vF1.1.01');
      Add('Titulo9=Registro de interdições e tutelas');
    end;

  StringList.SaveToFile(ExtractFilePath(Application.ExeName) + gsArqLocal);
  StringList.Free;
end;

procedure TFrmPrincipal.executarScripts;
var
   x : integer;
begin
  LimpaLabels;
  ModuloEmAtualizacao := 'Atualizando base de dados...';

  Application.ProcessMessages;

  idFtpAtualiza.ChangeDir( '/softwares/cartorio/SQL/' );
  idFtpAtualiza.List(LbArquivosRemotos.Items, '*.sql', false );

  if not SysUtils.DirectoryExists( ExtractFileDir(Application.ExeName ) + '\SQL' ) then
    CreateDir( ExtractFileDir(Application.ExeName ) + '\SQL' );

  conexaoBancoDados;

  GgProgressao.MaxValue := LbArquivosRemotos.Items.Count;
  GgProgressao.Progress := 0;

  for x := 0 to Pred(LbArquivosRemotos.Items.Count) do
    begin
       Application.ProcessMessages;
       if not FileExists( ExtractFileDir(Application.ExeName ) + '\SQL\' + LbArquivosRemotos.Items[x] ) then
         begin
           idFtpAtualiza.Get( '/softwares/cartorio/SQL/' + LbArquivosRemotos.Items[x],
             ExtractFileDir(Application.ExeName ) + '\SQL\' + LbArquivosRemotos.Items[x], true );

           MmScripts.Lines.Clear;
           MmScripts.Lines.LoadFromFile(ExtractFileDir(Application.ExeName ) + '\SQL\' + LbArquivosRemotos.Items[x]);;

           executaSQL(MmScripts.Lines.Text);
         end;

       GgProgressao.Progress := GgProgressao.Progress + 1;
       Application.ProcessMessages;
    end;

  GgProgressao.Progress := 0;
  Application.ProcessMessages;
end;

function TFrmPrincipal.executaSQL(sql: string) : boolean;
var
   Q : TSQLQuery;
begin
  result := true;

  try
    Q := TSQLQuery.Create(nil);
    Q.SQLConnection := SQLConnConexao;
    try
      Q.SQL.Clear;
      Q.SQL.Text := SQL;
      Q.ExecSQL();
    except
      result := false;
    end;
  finally
    FreeAndNil(Q);
  end;
end;

procedure TFrmPrincipal.IncluiDownloads(AModulo, AVersaoAtual,
  AVersaoRemota: string; ATamanho: LongWord);
begin
  with lstModulos.Items.Add do
    begin
      Caption := AModulo;
      SubItems.Append(AVersaoAtual);
      SubItems.Append(AVersaoRemota);
      SubItems.Append(FormatFloat('##,###,##0', ATamanho) + ' Kb');
    end;
end;

procedure TFrmPrincipal.LimpaLabels;
begin
  lblStatus.Caption := '';
  lblContador.Caption := '';
  lblNomeModulo.Caption := '';
  lblTotalBytes.Caption := '';
end;

procedure TFrmPrincipal.LimparPastaTemp;
var
  Lng          : DWORD;
  ThePath      : string;
  I            : integer;
  flbArquivos  : TFileListBox;
begin
  SetLength(thePath, MAX_PATH);
  Lng := GetTempPath(MAX_PATH, PChar(ThePath));
  SetLength(ThePath, Lng);
  flbArquivos := TFileListBox.Create(FrmPrincipal);
  flbArquivos.Parent := FrmPrincipal;
  flbArquivos.Directory := ThePath;
  flbArquivos.Visible := False;
  for I := flbArquivos.Items.Count - 1 downto 0 do
    DeleteFile(flbArquivos.Items[I]);
  flbArquivos.Free;
end;

procedure TFrmPrincipal.PosAtualizacao;
begin

end;

procedure TFrmPrincipal.PreencherModulosAbertos;
var
  I: Integer;
begin
  IniRemota := TIniFile.Create(ExtractFilePath(Application.ExeName) + gsArqRemoto);
  ModulosAbertos.Clear;
  for I := 1 to 20 do
  begin
    {Verifica se a chave existe, ou seja, se o "I" atual corresponde à algum módulo no arquivo. local.ini} 
    if not IniRemota.ValueExists('Atualizacoes', 'Modulo' + IntToStr(I)) then 
      Break; 
    {Nome e tamanho do Módulo} 
    Titulo := IniRemota.ReadString('Atualizacoes', 'Titulo' + IntToStr(I), Titulo); 
    if FindWindow(nil, PChar(Titulo)) > 0 then 
      ModulosAbertos.Add(Titulo); 
  end; 
  IniRemota.Free; 
end;

procedure TFrmPrincipal.TmVerificacaoTimer(Sender: TObject);
var 
  I : Integer;
begin
  try
    Screen.Cursor := crHourGlass;
    kbTotalDown := 0;
    Sleep(3000);
    lblListandoAtalizacoes.Caption := 'Verificando se há atualizações disponíveis';
    lblListandoAtalizacoes.Font.Color := clBlue;
    Update;

    with idFtpAtualiza do
    begin
      {Tenta conectar-se ao servidor FTP}
      if Conectar then
      begin
        try
          IniLocal := TIniFile.Create(ExtractFilePath(Application.ExeName) + gsArqLocal);
          IniRemota := TIniFile.Create(ExtractFilePath(Application.ExeName) + gsArqRemoto);
          lstModulos.Items.Clear;
          {Faz o download do arquivo Remoto.ini pra comparar as versões}
          Get(sArqRemota, sArqLocal, True);

          {Faz as comparações necessárias e atualiza a lista de atualizações pendentes}
          for I := 1 to 50 do
          begin
            {Verifica se a chave existe, ou seja, se o "I" atual corresponde à algum módulo no arquivo. local.ini}
            if not IniRemota.ValueExists('Atualizacoes', 'Modulo' + IntToStr(I)) then
              Break;

            {Nome e tamanho do Módulo}
            Modulo := IniRemota.ReadString('Atualizacoes', 'Modulo' + IntToStr(I), Modulo);
            ModuloEmAtualizacao := Modulo;

            Tamanho_Arquivo := Size(gsPastaRemota + '/' + Modulo);

            VersaoLocal := IniLocal.ReadString('Atualizacoes', 'Versao' + IntToStr(I), VersaoLocal);
            VersaoRemota := IniRemota.ReadString('Atualizacoes', 'Versao' + IntToStr(I), VersaoRemota);

            {Verifica se há atualizações e então atualiza a lista de downloads}
            if FileExists(ExtractFilePath(Application.ExeName) + '\' + Modulo ) then
              begin
                if (VersaoLocal <> VersaoRemota) or not
                  (IniLocal.ValueExists('Atualizacoes', 'Modulo' + IntToStr(I))) then
                begin
                  IncluiDownloads(ModuloEmAtualizacao, VersaoLocal, VersaoRemota, Tamanho_Arquivo);
                  kbTotalDown := kbTotalDown + Tamanho_Arquivo;
                end;
              end;
          end;
          idFtpAtualiza.Disconnect;

          IniLocal.Free;
          IniRemota.Free;

          if lstModulos.Items.Count > 0 then
          begin
            lstModulos.Visible := True;
            btnAtualizar.Enabled := True;
            lblTotalBytes.Caption := 'Total de bytes à baixar: ' + FormatFloat('##,###,##0', kbTotalDown) + ' Kb';
          end
          else
          begin
            lstModulos.Visible := False;
            lblListandoAtalizacoes.Caption := 'Não há atualizações disponíveis';
            lblListandoAtalizacoes.Font.Color := clRed;
            DeleteFile(gsPastaLocal + gsArqRemoto);
          end;
        except on E: Exception do 
          begin 
            MessageDlg('Ocorreu um erro durante o processo.' + #13#13 + 'Mensagem original: ' + #13 + 
              E.Message + #13#13 + 'O aplicativo será fechado.', mtError, [mbOK], 0); 
            Halt; 
          end; 
        end; 
      end; 
    end; 
  finally 
    Screen.Cursor := crDefault; 
    TmVerificacao.Enabled := False;
  end; 
end;

procedure TFrmPrincipal.btnAtualizarClick(Sender: TObject);
var
  I                 : Integer;
  iContaAtualizacao : Integer;
  localSistema      : string;
begin
  try
    Screen.Cursor := crHourGlass;
    {Tenta conectar-se ao servidor FTP da América}
    if Conectar then
    begin
      with idFtpAtualiza do
      begin
        IniLocal := TIniFile.Create(ExtractFilePath(Application.ExeName) + gsArqLocal);
        IniRemota := TIniFile.Create(ExtractFilePath(Application.ExeName) + gsArqRemoto);
        {Re-ativa os eventos do TidFtp}

        OnWorkBegin := idFtpAtualizaWorkBegin;
        OnWork := idFtpAtualizaWork;
        flagAtivado := true;
        try
          btnAtualizar.Enabled := False;
          btnFechar.Enabled := False;
          Tamanho_Arquivo := 1;

          IniLocal := TIniFile.Create(ExtractFilePath(Application.ExeName) + gsArqLocal);
          IniRemota := TIniFile.Create(ExtractFilePath(Application.ExeName) + gsArqRemoto);
          {Aqui verifica se há atualizações e baixa o módulo}
          iContaAtualizacao := 0;
          for I := 1 to 20 do
          begin
            {Verifica se a chave existe, ou seja, se o "I" atual corresponde à algum módulo no arquivo. local.ini}
            if not IniRemota.ValueExists('Atualizacoes', 'Modulo' + IntToStr(I)) then
              Break;

            {Nome e tamanho do Módulo}
            Modulo := IniRemota.ReadString('Atualizacoes', 'Modulo' + IntToStr(I), Modulo);
            titulo := IniRemota.ReadString('Atualizacoes', 'Titulo' + IntToStr(I), titulo);

            ModuloEmAtualizacao := titulo;
            Tamanho_Arquivo := Size(gsPastaRemota + '/' + Modulo);

            VersaoLocal := IniLocal.ReadString('Atualizacoes', 'Versao' + IntToStr(I), VersaoLocal);
            VersaoRemota := IniRemota.ReadString('Atualizacoes', 'Versao' + IntToStr(I), VersaoRemota);

            {Efetua o download}
            if (VersaoLocal <> VersaoRemota) or not
              (IniLocal.ValueExists('Atualizacoes', 'Modulo' + IntToStr(I))) then
            begin
              Inc(iContaAtualizacao);
              sArqRemota := gsPastaRemota + '/' + Modulo;
              sArqLocal := gsPastaLocal + '\' + Modulo;
              Tamanho_Arquivo := Size(sArqRemota);

              RenameFile( sArqLocal, Concat(ExtractFilePath(Application.ExeName),
                Copy( Modulo, 1, Pos( '.', Modulo ) -1 ), '_Data_Atualização=',
                  FormatDateTime( 'dd-mm-yyyy-hh-mm-ss', Now), '.exe' ) );

              Get(sArqRemota, sArqLocal, True);
              lstModulos.Items[iContaAtualizacao - 1].Checked := True;
              AtualizarINIVersoes('Atualizacoes', 'Versao' + IntToStr(I), VersaoRemota);
            end;
          end;

          //EXECUTANDO SCRIPTS DO BANCO DE DADOS
          executarScripts;

          Disconnect;

          IniLocal.Free;
          IniRemota.Free;

          {Função desativada - Atualizando a cada download}
          DeleteFile(gsPastaLocal + gsArqRemoto);

          lblListandoAtalizacoes.Caption := 'Todos os módulos listados foram atualizados com sucesso';
          lblListandoAtalizacoes.Font.Color := clBlue;
          btnAtualizar.Enabled := False;

          lstModulos.Visible := False;
          GgProgressao.Visible := false;

          btnFechar.Enabled := True; 
          LimpaLabels; 
          PosAtualizacao;
        except on E: Exception do 
          begin 
            MessageDlg('Ocorreu um erro durante o processo.' + #13#13 +
              E.Message, mtError, [mbOK], 0);
            Halt; 
          end; 
        end; 
      end; 
    end; 
  finally 
    Screen.Cursor := crDefault; 
  end; 
end;

procedure TFrmPrincipal.idFtpAtualizaWork(ASender: TObject;
  AWorkMode: TWorkMode; AWorkCount: Int64);
var
  Contador, kbTotal,
    kbTransmitidos,
      kbFaltantes   : Integer;
  Status_transf     : string;
  TotalTempo        : TDateTime;
  H, M, Sec, MS     : Word;
  DLTime, Media     : Double;
begin
  if flagAtivado then
    begin
      btnAtualizar.Enabled := False;
      btnFechar.Enabled := False;
      kbTotal := Tamanho_Arquivo div 1024;
      TotalTempo := Now - STime;
      DecodeTime(TotalTempo, H, M, Sec, MS);
      Sec := Sec + M * 60 + H * 3600;
      DLTime := Sec + MS / 1000;
      KbTransmitidos := AWorkCount div 1024;
      kbFaltantes := kbTotal - kbTransmitidos;
      lblContador.Caption := 'Transmitidos: ' + FormatFloat('##,###,##0', kbTransmitidos) +
        ' Kb de ' + FormatFloat('##,###,##0', kbTotal) + ' Kb' + '; Restam: ' + FormatFloat('##,###,##0', kbFaltantes) + ' Kb';
      Media := (100 / Tamanho_Arquivo) * AWorkCount;

      if DLTime > 0 then
      begin
        Tempo_Medio := (AWorkCount / 1024) / DLTime;
        Status_Transf := Format('%2d:%2d:%2d:', [Sec div 3600, (Sec div 60) mod 60, Sec mod 60]);
        Status_Transf := 'Tempo de download ' + Status_Transf;
      end;

      Status_Transf := 'Taxa de tranferência: ' +
        FormatFloat('0.00 Kb/s', Tempo_Medio) + '; ' + Status_Transf;
      lblStatus.Caption := Status_Transf;
      lblNomeModulo.Caption := ModuloEmAtualizacao;
      Application.ProcessMessages;
      Contador := Trunc(Media);
      GgProgressao.Progress := (contador);
    end;
end;

procedure TFrmPrincipal.idFtpAtualizaWorkBegin(ASender: TObject;
  AWorkMode: TWorkMode; AWorkCountMax: Int64);
begin
  if flagAtivado then
    begin
      STime := Now;
      Tempo_Medio := 0;
      pnlNomeModulo.Visible := True;
      Update;
    end;
end;

procedure TFrmPrincipal.FormCreate(Sender: TObject);
begin
  ApagarTemporarios;
  LimparPastaTemp;

  if not FileExists(ExtractFilePath(Application.ExeName) + gsArqLocal ) then
    CriaArquivoVersoes;

  LimpaLabels;
  TmVerificacao.Enabled := true;
end;

procedure TFrmPrincipal.GetValorArquivoIni(Secao, NomeVariavel: string;
  var Resultado: string);
var
   ArqIni : TIniFile;

   vString  : String;
begin
  vString  := '';

  try
    ArqIni := TIniFile.Create( ExtractFilePath( Application.ExeName ) + '\config.ini' );

    vString := ArqIni.ReadString( Secao, NomeVariavel , vString );
    Resultado := vString;
  finally
    FreeAndNil( ArqIni );
  end;
end;

end.
