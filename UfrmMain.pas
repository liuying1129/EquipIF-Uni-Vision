unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ToolWin, Inifiles, DB, ADODB, Grids, DBGrids, Menus,
  StdCtrls, ExtCtrls,PerlRegEx, DBCtrls, Buttons, ActnList, StrUtils;

type
  TfrmMain = class(TForm)
    ADOConnPEIS: TADOConnection;
    ADOConnEquip: TADOConnection;
    DataSource1: TDataSource;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    DataSource2: TDataSource;
    DataSource3: TDataSource;
    ADOQuery3: TADOQuery;
    StatusBar1: TStatusBar;
    GroupBox1: TGroupBox;
    DBGrid2: TDBGrid;
    PageControl2: TPageControl;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    DBGrid3: TDBGrid;
    Panel2: TPanel;
    Label2: TLabel;
    Memo2: TMemo;
    DBNavigator2: TDBNavigator;
    GroupBox2: TGroupBox;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    DBGrid1: TDBGrid;
    TabSheet2: TTabSheet;
    Memo1: TMemo;
    Panel1: TPanel;
    DBNavigator1: TDBNavigator;
    BitBtn1: TBitBtn;
    ActionList1: TActionList;
    Action1: TAction;
    Label1: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    DateTimePicker1: TDateTimePicker;
    Label6: TLabel;
    Label7: TLabel;
    LabeledEdit1: TLabeledEdit;
    DateTimePicker2: TDateTimePicker;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ADOQuery1AfterOpen(DataSet: TDataSet);
    procedure ADOQuery1AfterScroll(DataSet: TDataSet);
    procedure ADOQuery2AfterScroll(DataSet: TDataSet);
    procedure ADOQuery2AfterOpen(DataSet: TDataSet);
    procedure ADOQuery3AfterOpen(DataSet: TDataSet);
    procedure ADOQuery3AfterScroll(DataSet: TDataSet);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DateTimePicker1Change(Sender: TObject);
    procedure DateTimePicker2Change(Sender: TObject);
    procedure LabeledEdit1KeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
  private
    { Private declarations }
    function MakeDBConn(const ADB:string;AADOConn:TADOConnection):boolean;
    procedure UpdateEquipAdoquery;
    function BuildPeisSQL(const AName,ASex,AAge:string):String;
    procedure UpdateAdoquery3;
    procedure GetEquipJcts(Sender: TField; var Text: String;DisplayText: Boolean);
    function singleSend2Peis(const AEquipUnid,AEquipName,AEquipSex,AEquipAge,AEquipJcts:String):boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses UCommFunction;

const
  sCryptSeed='lc';//�ӽ�������
  SYSNAME='PEIS'; 

var
  PeisConnStr:String;
  EquipConnStr:String;
  ArCheckBoxValue:TStrings;//keyΨһ,���ܱ�֤ҵ���߼���ȷ

{$R *.dfm}

function GetConnectString(const ADB:string):string;
//ADB:PEIS���ݿ⡢�豸���ݿ�
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//�Ƿ񼯳ɵ�¼ģʽ

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString(ADB, '������', '');
  initialcatalog := Ini.ReadString(ADB, '���ݿ�', '');
  ifIntegrated:=ini.ReadBool(ADB,'���ɵ�¼ģʽ',false);
  userid := Ini.ReadString(ADB, '�û�', '');
  password := Ini.ReadString(ADB, '����', '107DFC967CDCFAAF');
  Ini.Free;
  //======����password
  pInStr:=pchar(password);
  pDeStr:=DeCryptStr(pInStr,sCryptSeed);
  setlength(password,length(pDeStr));
  for i :=1  to length(pDeStr) do password[i]:=pDeStr[i-1];
  //==========

  result := result + 'user id=' + UserID + ';';
  result := result + 'password=' + Password + ';';
  result := result + 'data source=' + datasource + ';';
  result := result + 'Initial Catalog=' + initialcatalog + ';';
  result := result + 'provider=' + 'SQLOLEDB.1' + ';';
  //Persist Security Info,��ʾADO�����ݿ����ӳɹ����Ƿ񱣴�������Ϣ
  //ADOȱʡΪTrue,ADO.netȱʡΪFalse
  //�����лᴫADOConnection��Ϣ��TADOLYQuery,������ΪTrue
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

function TfrmMain.MakeDBConn(const ADB:string;AADOConn:TADOConnection): boolean;
//ADB:PEIS���ݿ⡢�豸���ݿ�
var
  newconnstr,ss: string;
  Label labReadIni;
begin
  result:=false;

  labReadIni:
  newconnstr := GetConnectString(ADB);
  
  try
    AADOConn.Connected := false;
    AADOConn.ConnectionString := newconnstr;
    AADOConn.Connected := true;
    result:=true;
    if ADB='PEIS���ݿ�' then PeisConnStr:=newconnstr;
    if ADB='�豸���ݿ�' then EquipConnStr:=newconnstr;
  except
  end;
  if not result then
  begin
    ss:='������'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '���ݿ�'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '���ɵ�¼ģʽ'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '�û�'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '����'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('�������ݿ�',Pchar(ADB),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.UpdateEquipAdoquery;
var
  ss1:string;
begin
  ss1:='SELECT TOP 3000 '+//2���µ������
        'P.patient_name as ���� '+
       ',P.patient_age as ���� '+
       ',P.patient_sex as �Ա� '+
       ',0 as ѡ�� '+
       ',S.study_dttm as ����ʱ�� '+//X��ϵͳ����ļ��ʱ��
       ',S.deft_name as �ͼ���� '+
       ',S.refer_doctor as �ͼ�ҽ�� '+
  	   ',R.Diagnosis as �����ʾ '+
  	   ',P.patient_id '+
	     ',R.Finding as ������� '+
	     ',P.patient_key '+
       ',P.patient_birth_date '+
       ',S.study_key '+
       ',S.study_instance_uid '+
       ',S.study_id '+
       ',S.study_desc '+
       ',S.series_count '+
       ',S.instance_count '+
       ',S.modality '+
       ',S.bodypart '+
       ',S.status '+
       ',S.bed_no '+
       ',S.inp_no '+
       ',S.approver '+
       ',S.approval_dttm '+
       ',S.isabnoraml '+
       ',S.costs '+
       ',S.patsource '+
	     ',S.writeingreported '+
	     ',R.report_key '+
	     ',R.transcriber '+
	     ',R.approver '+
	     ',R.approval_dttm '+
      ',PS.SendSuccNum '+// AS ���ͳɹ�����
      //',PS.LastSendDes AS �������� '+
      //',PS.LastSendTime AS �����ʱ�� '+
' FROM Patient P '+
'LEFT JOIN Study S ON S.patient_key=P.patient_key '+
'LEFT JOIN Report R ON R.study_key=S.study_key '+
  'LEFT JOIN PEIS_Send PS ON PS.StudyResultIdentity=R.report_key '+
  ' where S.status in (''�ѱ���'',''�����'') '+//ֻ��ѯ����ɱ��浥
  ' and S.deft_name=''���'' '+//ֻ��ѯ��챨�浥(�������=���)
  ' and P.patient_name is not null and P.patient_name<>'''' '+//������������
  ' and datalength(R.Diagnosis)>0 '+//�޼����ʾ������//��ζ��report_key��ֵ
  ' AND S.study_dttm between :begin_study_dttm and :end_study_dttm '+
  ifThen(trim(LabeledEdit1.Text)<>'',' and P.patient_id like '''+trim(LabeledEdit1.Text)+'%'' ')+
  ' order by p.patient_key desc';

//Patient:������Ϣ��
//Study:�ͼ�ҽ�����ͼ���ҡ����ҽ������鲿λ�����
//Report:��������ٴ����
//ImageRecord:DCOMͼƬ
//Reportpicture:ͼƬ

  ADOQuery1.Close;
  ADOQuery1.SQL.Clear;
  ADOQuery1.SQL.Add(ss1);
  ADOQuery1.Parameters.ParamByName('begin_study_dttm').Value:=DateTimePicker1.DateTime;
  ADOQuery1.Parameters.ParamByName('end_study_dttm').Value:=DateTimePicker2.DateTime;
  ADOQuery1.Open;
end;

procedure TfrmMain.FormShow(Sender: TObject);
begin
  StatusBar1.Panels[2].Text:=SYSNAME;

  DateTimePicker1.Date := now-30;
  DateTimePicker1.Time := StrToTime('00:00:01');
  DateTimePicker2.Date := now;
  DateTimePicker2.Time := StrToTime('23:59:59');

  UpdateEquipAdoquery;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  MakeDBConn('�豸���ݿ�',ADOConnEquip);
  MakeDBConn('PEIS���ݿ�',ADOConnPEIS);

  PageControl1.ActivePageIndex:=0;
  PageControl2.ActivePageIndex:=0;
end;

procedure TfrmMain.GetEquipJcts(Sender: TField; var Text: String; DisplayText: Boolean);
begin
  Text:=Sender.AsString;
end;

procedure TfrmMain.ADOQuery1AfterOpen(DataSet: TDataSet);
begin
  dbgrid1.Columns[0].Width:=42;
  dbgrid1.Columns[1].Width:=30;
  dbgrid1.Columns[2].Width:=30;
  dbgrid1.Columns[3].Width:=30;//ѡ��
  dbgrid1.Columns[4].Width:=72;
  dbgrid1.Columns[5].Width:=60;
  dbgrid1.Columns[6].Width:=55;
  dbgrid1.Columns[7].Width:=300;
  dbgrid1.Columns[8].Width:=73;

  DataSet.FieldByName('�����ʾ').OnGetText:=GetEquipJcts;
  
  ArCheckBoxValue.Clear;
end;

function TfrmMain.singleSend2Peis(const AEquipUnid,AEquipName,AEquipSex,AEquipAge,AEquipJcts:String): boolean;
var
  adotemp3,adotemp4,adotemp5,adotemp555:tadoquery;
  jcts_itemid:string;//�������ʾ����Ŀ����
  jcjl_itemid:string;//�������ۡ���Ŀ����
  jcjy_itemid:string;//����齨�顿��Ŀ����
  jcjl_combinid:string;//�������ۡ������Ŀ����
  jcjy_combinid:string;//����齨�顿�����Ŀ����
  Eqip_Jcts2:String;
  Eqip_Jcts4:String;
  Peis_Unid:String;
  Eqip_Jcts_List:TStrings;
  RegEx: TPerlRegEx;
  RegEx2: TPerlRegEx;
  RegEx3: TPerlRegEx;
  RegEx4: TPerlRegEx;
  b3:boolean;
  i:integer;

  Peis_Jcjl:String;
  Peis_Jcjy:String;

  Peis_Jcts_Num:integer;
  Peis_Jcjl_Num:integer;
  Peis_Jcjy_Num:integer;

  adotemp44:tadoquery;
begin
  result:=false;

  adotemp44:=tadoquery.Create(nil);
  adotemp44.Connection:=ADOConnPEIS;
  adotemp44.Close;
  adotemp44.SQL.Clear;
  adotemp44.SQL.Text:=BuildPeisSQL(AEquipName,AEquipSex,AEquipAge);
  adotemp44.Open;
  if adotemp44.RecordCount<=0 then
  begin
    MESSAGEDLG(AEquipName+':PEIS�޸��ܼ���!',mtError,[MBOK],0);
    adotemp44.Free;
    exit;
  end;
  if adotemp44.RecordCount>1 then
  begin
    MESSAGEDLG(AEquipName+':���ܼ�����PEIS�ж�����¼!',mtError,[MBOK],0);
    adotemp44.Free;
    exit;
  end;
  if adotemp44.FieldByName('�����').AsString<>'' then
  begin
    MESSAGEDLG(AEquipName+':���ܼ�����PEIS�����!',mtError,[MBOK],0);
    adotemp44.Free;
    exit;
  end;
  Peis_Unid:=adotemp44.fieldbyname('Unid').AsString;
  adotemp44.Free;

  adotemp3:=tadoquery.Create(nil);
  adotemp3.Connection:=ADOConnPEIS;
  adotemp3.Close;
  adotemp3.SQL.Clear;
  adotemp3.SQL.Text:='select itemid from clinicchkitem cci where cci.Reserve5=4 ';
  adotemp3.Open;
  if adotemp3.RecordCount<=0 then
  begin
    MESSAGEDLG('������X�⡾�����ʾ����Ŀ(�����ֶ�5=4).�����Ա�����Ŀ����!',mtError,[MBOK],0);
    adotemp3.Free;
    exit;
  end;
  if adotemp3.RecordCount>1 then
  begin
    MESSAGEDLG('���ڶ���X�⡾�����ʾ����Ŀ(�����ֶ�5=4).�����Ա�����Ŀ����!',mtError,[MBOK],0);
    adotemp3.Free;
    exit;
  end;
  jcts_itemid:=adotemp3.fieldbyname('itemid').AsString;
  adotemp3.Free;

  adotemp4:=tadoquery.Create(nil);
  adotemp4.Connection:=ADOConnPEIS;
  adotemp4.Close;
  adotemp4.SQL.Clear;
  adotemp4.SQL.Text:='select cci.itemid,cbi.id from clinicchkitem cci,CombSChkItem csi,combinitem cbi where cci.Reserve5=1 and cci.unid=csi.itemunid and csi.combunid=cbi.unid and cci.SysName=''PEIS'' and cbi.SysName=''PEIS'' ';
  adotemp4.Open;
  if adotemp4.RecordCount<=0 then
  begin
    MESSAGEDLG('�����ڡ������ۡ���Ŀ(�����ֶ�5=1)�򲻴��������Ŀ.�����Ա�����Ŀ����!',mtError,[MBOK],0);
    adotemp4.Free;
    exit;
  end;
  if adotemp4.RecordCount>1 then
  begin
    MESSAGEDLG('���ڶ����������ۡ���Ŀ(�����ֶ�5=1)�����ڶ�������Ŀ.�����Ա�����Ŀ����!',mtError,[MBOK],0);
    adotemp4.Free;
    exit;
  end;
  jcjl_itemid:=adotemp4.fieldbyname('itemid').AsString;
  jcjl_combinid:=adotemp4.fieldbyname('id').AsString;
  adotemp4.Free;

  adotemp5:=tadoquery.Create(nil);
  adotemp5.Connection:=ADOConnPEIS;
  adotemp5.Close;
  adotemp5.SQL.Clear;
  adotemp5.SQL.Text:='select cci.itemid,cbi.id from clinicchkitem cci,CombSChkItem csi,combinitem cbi where cci.Reserve5=2 and cci.unid=csi.itemunid and csi.combunid=cbi.unid and cci.SysName=''PEIS'' and cbi.SysName=''PEIS'' ';
  adotemp5.Open;
  if adotemp5.RecordCount<=0 then
  begin
    MESSAGEDLG('�����ڡ���齨�顿��Ŀ(�����ֶ�5=2)�򲻴��������Ŀ.�����Ա�����Ŀ����!',mtError,[MBOK],0);
    adotemp5.Free;
    exit;
  end;
  if adotemp5.RecordCount>1 then
  begin
    MESSAGEDLG('���ڶ�������齨�顿��Ŀ(�����ֶ�5=2)�����ڶ�������Ŀ.�����Ա�����Ŀ����!',mtError,[MBOK],0);
    adotemp5.Free;
    exit;
  end;
  jcjy_itemid:=adotemp5.fieldbyname('itemid').AsString;
  jcjy_combinid:=adotemp5.fieldbyname('id').AsString;
  adotemp5.Free;

  Peis_Jcts_Num:=strtoint(ScalarSQLCmd(PeisConnStr,'select count(*) from chk_valu cv where cv.pkunid='+Peis_Unid+' and cv.itemid='''+jcts_itemid+''' '));
  if Peis_Jcts_Num<=0 then
  begin
    ExecSQLCmd(PeisConnStr,'insert into chk_valu (pkunid,itemid,itemvalue) values ('+Peis_Unid+','''+jcts_itemid+''','''+AEquipJcts+''')');
  end else
  begin
    //if (MessageDlg('PEIS���ڼ������,������ԭ�м������,ȷ����', mtConfirmation, [mbYes, mbNo], 0) <> mrYes) then exit;
    ExecSQLCmd(PeisConnStr,'update chk_valu set itemvalue='''+AEquipJcts+''' where pkunid='+Peis_Unid+' and itemid='''+jcts_itemid+''' ');
  end;

  RegEx := TPerlRegEx.Create;
  RegEx.Subject := AEquipJcts;
  RegEx.RegEx   := '��|��';//��|�ָ�����ָ���.������ʽ|��ʾ"��"
  Eqip_Jcts_List:=TStringList.Create;
  try
    RegEx.Split(Eqip_Jcts_List,MaxInt);//MaxInt,��ʾ�ֶܷ��پͷֶ���
  except
    on E:Exception do
    begin
      WriteLog(pchar(AEquipName+':RegEx.Splitʧ��:'+E.Message));
      MESSAGEDLG(AEquipName+':RegEx.Splitʧ��:'+E.Message,mtError,[mbOK],0);
      FreeAndNil(RegEx);
      Eqip_Jcts_List.Free;
      exit;
    end;
  end;
  FreeAndNil(RegEx);
  for i :=0  to Eqip_Jcts_List.Count-1 do
  begin
    if pos('δ��ʵ���Բ���',Eqip_Jcts_List[i])>0 then continue;
    if pos('δ����Բ���',Eqip_Jcts_List[i])>0 then continue;
    if pos('δ���쳣',Eqip_Jcts_List[i])>0 then continue;//��׵����δ���쳣

    //���ɼ�����begin
    //ɾ�������ʾ�е����(��1��23��1.23.)
    RegEx2 := TPerlRegEx.Create;
    RegEx2.Subject := trim(Eqip_Jcts_List[i]);
    RegEx2.RegEx   := '^\d{1,2}(��|\.)';//.Ϊ������ʽ��Ԫ�ַ�������\ת��
    RegEx2.Replacement:='';
    try
      RegEx2.ReplaceAll;
    except
      on E:Exception do
      begin
        WriteLog(pchar(AEquipName+':Subject:'+trim(Eqip_Jcts_List[i])+'.ɾ�������ʾ���RegEx.ReplaceAllʧ��:'+E.Message));
        MESSAGEDLG(AEquipName+':Subject:'+trim(Eqip_Jcts_List[i])+'.ɾ�������ʾ���RegEx.ReplaceAllʧ��:'+E.Message,mtError,[mbOK],0);
        FreeAndNil(RegEx2);
        Eqip_Jcts_List.Free;
        exit;
      end;
    end;
    Eqip_Jcts2:=RegEx2.Subject;
    FreeAndNil(RegEx2);

    //ɾ�������ʾ�еĽ���
    RegEx4 := TPerlRegEx.Create;
    RegEx4.Subject := Eqip_Jcts2;
    RegEx4.RegEx   := '������[\s\S]*';
    RegEx4.Replacement:='';
    try
      RegEx4.ReplaceAll;
    except
      on E:Exception do
      begin
        WriteLog(pchar(AEquipName+':Subject:'+Eqip_Jcts2+'.ɾ�������ʾ����RegEx.ReplaceAllʧ��:'+E.Message));
        MESSAGEDLG(AEquipName+':Subject:'+Eqip_Jcts2+'.ɾ�������ʾ����RegEx.ReplaceAllʧ��:'+E.Message,mtError,[mbOK],0);
        FreeAndNil(RegEx4);
        Eqip_Jcts_List.Free;
        exit;
      end;
    end;
    Eqip_Jcts4:=RegEx4.Subject;
    FreeAndNil(RegEx4);

    if trim(Eqip_Jcts4)<>'' then Peis_Jcjl:=Peis_Jcjl+trim(Eqip_Jcts4)+'��'+#13;//������
    //���ɼ�����end

    //���ɼ�齨��begin
    adotemp555:=tadoquery.Create(nil);
    adotemp555.Connection:=ADOConnPEIS;
    adotemp555.Close;
    adotemp555.SQL.Clear;
    adotemp555.SQL.Text:='select name,Reserve2 from CommCode where TypeName=''�쳣����'' ';
    adotemp555.Open;
    while not adotemp555.Eof do
    begin
      //ƥ���쳣�ؼ���
      RegEx3 := TPerlRegEx.Create;
      RegEx3.Subject := Eqip_Jcts_List[i];
      RegEx3.RegEx   := adotemp555.fieldbyname('name').AsString;
      try
        b3:=RegEx3.Match;
      except
        on E:Exception do
        begin
          WriteLog(pchar(AEquipName+':Subject:'+Eqip_Jcts_List[i]+'.RegEx.Matchʧ��:'+E.Message+'.������ʽ:'+adotemp555.fieldbyname('name').AsString));
          MESSAGEDLG(AEquipName+':Subject:'+Eqip_Jcts_List[i]+'.RegEx.Matchʧ��:'+E.Message+'.������ʽ:'+adotemp555.fieldbyname('name').AsString,mtError,[mbOK],0);
          FreeAndNil(RegEx3);
          Eqip_Jcts_List.Free;
          adotemp555.Free;
          exit;
        end;
      end;
      FreeAndNil(RegEx3);
      if b3 then Peis_Jcjy:=Peis_Jcjy+adotemp555.fieldbyname('Reserve2').AsString+#13;
      adotemp555.Next;
    end;
    adotemp555.Free;
    //���ɼ�齨��end
  end;
  Eqip_Jcts_List.Free;

  Peis_Jcjl_Num:=strtoint(ScalarSQLCmd(PeisConnStr,'select count(*) from chk_valu cv where cv.pkunid='+Peis_Unid+' and cv.itemid='''+jcjl_itemid+''' '));
  if Peis_Jcjl_Num<=0 then
  begin
    ExecSQLCmd(PeisConnStr,'insert into chk_valu (pkunid,pkcombin_id,itemid,itemvalue) values ('+Peis_Unid+','''+jcjl_combinid+''','''+jcjl_itemid+''','''+Peis_Jcjl+''')');
  end else
  begin
    ExecSQLCmd(PeisConnStr,'update chk_valu set itemvalue=itemvalue+'''+Peis_Jcjl+''' where pkunid='+Peis_Unid+' and itemid='''+jcjl_itemid+''' ');
  end;

  Peis_Jcjy_Num:=strtoint(ScalarSQLCmd(PeisConnStr,'select count(*) from chk_valu cv where cv.pkunid='+Peis_Unid+' and cv.itemid='''+jcjy_itemid+''' '));
  if Peis_Jcjy_Num<=0 then
  begin
    ExecSQLCmd(PeisConnStr,'insert into chk_valu (pkunid,pkcombin_id,itemid,itemvalue) values ('+Peis_Unid+','''+jcjy_combinid+''','''+jcjy_itemid+''','''+Peis_Jcjy+''')');
  end else
  begin
    ExecSQLCmd(PeisConnStr,'update chk_valu set itemvalue=itemvalue+'''+Peis_Jcjy+''' where pkunid='+Peis_Unid+' and itemid='''+jcjy_itemid+''' ');
  end;

  if strtoint(ScalarSQLCmd(EquipConnStr,'select count(*) from PEIS_Send ps where ps.StudyResultIdentity='+AEquipUnid))<=0 then
    ExecSQLCmd(EquipConnStr,'insert into PEIS_Send (StudyResultIdentity,SendSuccNum) values ('+AEquipUnid+',1)')
  else ExecSQLCmd(EquipConnStr,'update PEIS_Send set SendSuccNum=SendSuccNum+1 where StudyResultIdentity='+AEquipUnid);

  result:=true;
end;

procedure TfrmMain.ADOQuery1AfterScroll(DataSet: TDataSet);
begin
    ADOQuery2.Close;
    ADOQuery2.SQL.Clear;
    ADOQuery2.SQL.Text:=BuildPeisSQL(DataSet.FieldByName('����').AsString,DataSet.FieldByName('�Ա�').AsString,DataSet.FieldByName('����').AsString);
    ADOQuery2.Open;

  //���½������
  Memo1.Lines.Text:=DataSet.fieldbyname('����').AsString+
              ' '+
              DataSet.fieldbyname('����').AsString+
              ' '+
              DataSet.fieldbyname('�Ա�').AsString+
              ' '+
              FormatDateTime('YYYY-MM-DD',DataSet.fieldbyname('����ʱ��').AsDateTime)+
              #13#13+
              DataSet.fieldbyname('�����ʾ').AsString;
end;

procedure TfrmMain.ADOQuery2AfterScroll(DataSet: TDataSet);
begin
  UpdateAdoquery3;
end;

procedure TfrmMain.ADOQuery2AfterOpen(DataSet: TDataSet);
begin
  dbgrid2.Columns[0].Width:=42;
  dbgrid2.Columns[1].Width:=30;
  dbgrid2.Columns[2].Width:=30;
  dbgrid2.Columns[3].Width:=72;
  dbgrid2.Columns[4].Width:=42;

  if DataSet.RecordCount<1 then UpdateAdoquery3;
end;

procedure TfrmMain.ADOQuery3AfterOpen(DataSet: TDataSet);
begin
  dbgrid3.Columns[1].Width:=80;
  dbgrid3.Columns[2].Width:=50;
  dbgrid3.Columns[3].Width:=300;
  dbgrid3.Columns[4].Width:=50;
  dbgrid3.Columns[5].Width:=50;
  dbgrid3.Columns[6].Width:=50;

  //���½������
  if DataSet.RecordCount<=0 then
  begin
    Memo2.Lines.Clear;
    Label2.Caption:='';
  end;
end;

procedure TfrmMain.UpdateAdoquery3;
var
  ss1:string;
  ss2:string;
begin
  if (ADOQuery2.Active)and(ADOQuery2.RecordCount>=1) then ss2:=ADOQuery2.fieldbyname('Unid').AsString else ss2:='-1';
  
    	ss1:=' select '+
      '(case cv.issure when 1 then ''��'' else '''' end) as ��,'+
      ' cv.Name as ����,cv.english_name as Ӣ����,cv.itemvalue as ���,cv.Min_value as ��Сֵ,cv.Max_value as ���ֵ,cv.Unit as ��λ,cv.pkcombin_id,cv.combin_Name,cv.pkunid,cv.valueid '+
    	' from chk_valu cv '+
    	' where '+
      ' cv.pkunid='+ss2+
      ' and cv.Reserve5 in (4,1,2) '+
      ' order by cv.printorder ';

    ADOQuery3.Close;
    ADOQuery3.SQL.Clear;
    ADOQuery3.SQL.Text:=SS1;
    ADOQuery3.Open;
end;

procedure TfrmMain.ADOQuery3AfterScroll(DataSet: TDataSet);
begin
  //���½������
  Memo2.Lines.Text:=DataSet.fieldbyname('���').AsString;
  Label2.Caption:=DataSet.fieldbyname('����').AsString;
end;

procedure TfrmMain.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
const
  CtrlState: array[Boolean] of Integer = (DFCS_BUTTONCHECK, DFCS_BUTTONCHECK or DFCS_CHECKED);
var
  strSendSuccNum:String;
  SendSuccNum:integer;
  
  checkBox_check:boolean;
begin
  //���͹��������б仯��ɫ
  if datacol=0 then
  begin
    strSendSuccNum:=tdbgrid(sender).DataSource.DataSet.fieldbyname('SendSuccNum').AsString;
    SendSuccNum:=strtointdef(strSendSuccNum,0);
    IF SendSuccNum>0 then
    begin
      tdbgrid(sender).Canvas.Font.Color:=clred;
      tdbgrid(sender).DefaultDrawColumnCell(rect,datacol,column,state);
    end;
  end;

  if Column.Field.FieldName='ѡ��' then
  begin
    (sender as TDBGrid).Canvas.FillRect(Rect);
    checkBox_check:=ArCheckBoxValue.Values[(Sender AS TDBGRID).DataSource.DataSet.FieldByName('report_key').AsString]='1';
    DrawFrameControl((sender as TDBGrid).Canvas.Handle,Rect, DFC_BUTTON, CtrlState[checkBox_check]);
  end else (sender as TDBGrid).DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;

procedure TfrmMain.DateTimePicker1Change(Sender: TObject);
begin
  UpdateEquipAdoquery;//ADOQuery1.Requery;
end;

procedure TfrmMain.DateTimePicker2Change(Sender: TObject);
begin
  UpdateEquipAdoquery;//ADOQuery1.Requery;
end;

procedure TfrmMain.LabeledEdit1KeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key<>13 then EXIT;

  UpdateEquipAdoquery;//ADOQuery1.Requery;

  if (Sender as TLabeledEdit).CanFocus then begin (Sender as TLabeledEdit).SetFocus;(Sender as TLabeledEdit).SelectAll; end;
end;

procedure TfrmMain.DBGrid1CellClick(Column: TColumn);
begin
  if not Column.Grid.DataSource.DataSet.Active then exit;  
  if Column.Field.FieldName <>'ѡ��' then exit;

  //TStringList���޸ü���������Զ�����
  ArCheckBoxValue.Values[Column.Grid.DataSource.DataSet.FieldByName('report_key').AsString]:=
    ifThen(ArCheckBoxValue.Values[Column.Grid.DataSource.DataSet.FieldByName('report_key').AsString]='1','0','1');
  Column.Grid.Refresh;//����DBGrid1DrawColumnCell�¼�
end;

procedure TfrmMain.SpeedButton1Click(Sender: TObject);
var
  i:integer;
begin
  for i:=0 to ArCheckBoxValue.Count-1 do
  begin
    ArCheckBoxValue.ValueFromIndex[i]:='1';
  end;
  DBGrid1.Refresh;//����DBGrid1DrawColumnCell�¼�
end;

procedure TfrmMain.SpeedButton2Click(Sender: TObject);
var
  i:integer;
begin
  for i:=0 to ArCheckBoxValue.Count-1 do
  begin
    ArCheckBoxValue.ValueFromIndex[i]:='0';
  end;
  DBGrid1.Refresh;//����DBGrid1DrawColumnCell�¼�
end;

function TfrmMain.BuildPeisSQL(const AName, ASex, AAge: string): String;
begin
  result:= 'select cc.patientname as ����,cc.age as ����,cc.sex as �Ա�,cc.check_date as �������,cc.report_doctor as �����,cc.combin_id as ������,cc.Caseno as ������,cc.deptname as �ͼ����,cc.check_doctor as �ͼ�ҽ��,cc.unid '+
    	' from chk_con cc,CommCode cco '+
    	' where cco.TypeName=''�������'' and cco.SysName='''+SYSNAME+
      ''' and cc.combin_id=cco.name '+
      ' AND cc.patientname='''+AName+
    	''' and isnull(cc.sex,'''')='''+ASex+
    	''' and dbo.uf_GetAgeReal(cc.age)=dbo.uf_GetAgeReal('''+AAge+''') ';
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
var
  Save_Cursor:TCursor;
  adotemp11:tadoquery;
  ifSelect:boolean;
  StudyResultIdentity:integer;
  i:integer;
begin
  if (not ADOQuery1.Active)or(ADOQuery1.RecordCount<=0) then
  begin
    MESSAGEDLG('�޷�������!',mtError,[MBOK],0);
    exit;
  end;

  ifSelect:=false;
  for i:=0 to ArCheckBoxValue.Count-1 do
  begin
    if ArCheckBoxValue.ValueFromIndex[i]='1' then begin ifSelect:=true;break;end;
  end;
  if not ifSelect then begin MESSAGEDLG('δѡ��,�޷�������!',mtError,[MBOK],0);exit;end;

  (Sender as TBitBtn).Enabled:=false;

  Save_Cursor := Screen.Cursor;
  Screen.Cursor := crHourGlass;

  adotemp11:=tadoquery.Create(nil);
  adotemp11.clone(ADOQuery1);
  while not adotemp11.Eof do
  begin
    if ArCheckBoxValue.Values[adotemp11.fieldbyname('report_key').AsString]='1' then
      singleSend2Peis(adotemp11.FieldByName('report_key').AsString,adotemp11.FieldByName('����').AsString,adotemp11.FieldByName('�Ա�').AsString,adotemp11.FieldByName('����').AsString,adotemp11.FieldByName('�����ʾ').AsString);

    adotemp11.Next;
  end;
  adotemp11.Free;

  //����ˢ���ѷ��͵���ɫ//�����һ���ᴥ��UpdateAdoquery3�������������ٴ�ADOQuery3.Requery;
  StudyResultIdentity:=ADOQuery1.fieldbyname('report_key').AsInteger;
  UpdateEquipAdoquery;
  ADOQuery1.Locate('report_key',StudyResultIdentity,[loCaseInsensitive]) ;

  Screen.Cursor := Save_Cursor;  { Always restore to normal }
  
  (Sender as TBitBtn).Enabled:=true;

  MESSAGEDLG('���Ͳ��������!',mtInformation,[MBOK],0);
end;

initialization
  ArCheckBoxValue:=TStringList.Create;

finalization
  ArCheckBoxValue.Free;

end.
