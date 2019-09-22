Attribute VB_Name = "Module1"
Option Explicit

Sub CSV�Ǎ��}�N��()
    
    Dim File_Path As String
    Dim CSV_Array As Variant
    
    '-----------------------------------------------
    '�� �`�F�b�N��Ɓi�{�i�I�ɍ�Ƃ���O�Ɋm�F�j
    If Not FnGetTitleList(CSV_Array) Then Exit Sub  'CSV�̐ݒ���̓ǂݎ��B�G���[������΃}�N�����I������
    If Not FnFilePicker(File_Path) Then Exit Sub    '�t�@�C���̓ǂݍ��݁B�p�X���擾�ł��Ȃ��ꍇ�̓}�N�����I������
    
    
    '-----------------------------------------------
    '�� �������
    MacroMode = True        '���׌y��
    
    Call CSV_Initialaize
    
    '-----------------------------------------------
    '�� �f�[�^�̓ǂݎ��
    
    Call CSV_read_Macro(File_Path, CSV_Array)
    
    
    '-----------------------------------------------
    '�� �I������
    
    Worksheets("CSV").Activate
    MacroMode = False        '���׌y���@�߂�
    
    MsgBox "��Ƃ��I�����܂����B" & vbCrLf & _
            "�uCSV�v�V�[�g���m�F���Ă��������B"
    
End Sub


'====================================================================
'�@CSV�f�[�^�e��̐ݒ�l���擾����֐�
'�@�Ԃ�l�FTrue/False (Boolean�^)
'�@�@�@�@�@True �F�ݒ�l���擾�ł����ꍇ
'�@�@�@�@�@False�F�ݒ�l���擾�ł��Ȃ������ꍇ�i�����l�j
'�@�����@�FCSV_Array�@�ݒ�l�̔z��ilong�^��1�����z��j
'�@�@�@�@�@�@�@�@�@�@(ByRef�w��ɂ��)�擾�����ݒ�l���i�[���ĕԂ�
'====================================================================

Private Function FnGetTitleList(ByRef CSV_Array As Variant) As Boolean
    
    On Error GoTo Err_Data      '�r���ŕs���ȃG���[�����������ꍇ�A�G���[���b�Z�[�W��\��
    
    '-----------------------------------------------
    '�� �f�[�^�̓ǂݎ��
'    Dim sh As Worksheet
'    Set sh = Worksheets("�Ǎ��ݒ�")
    
    Dim �ݒ胊�X�g As Variant
    �ݒ胊�X�g = Worksheets("�Ǎ��ݒ�").Range("B6").CurrentRegion
    
    Dim ���� As Long
    ���� = UBound(�ݒ胊�X�g, 1) - 1        '�^�C�g���s������������
    
    If ���� < 1 Then            '������0���̏ꍇ�G���[����
        MsgBox "[�ǂݍ��ݐݒ�]���ݒ肳��Ă��܂���" & vbCrLf & _
                "�u�Ǎ��ݒ�v�V�[�g�̐ݒ�����Ă�������", _
                vbCritical, _
                "�u�Ǎ��ݒ�v�V�[�g ����0�@�G���["
        Exit Function
    End If
    
    '-----------------------------------------------
    '�� �f�[�^�̏����o������
    Dim i As Long
    Dim tmp() As Long
    ReDim tmp(���� - 1)
    
    For i = 1 To ����
        tmp(i - 1) = CLng(Left(�ݒ胊�X�g(i + 1, 2), 1))
    Next i
    
    CSV_Array = tmp         '�z��̊i�[
    FnGetTitleList = True   '�������ʂ�Ԃ�
    
    
    Exit Function           '�����̏I��
    
Err_Data:
    MsgBox "[�ǂݍ��ݐݒ�]��ǂݎ�莞�ɕs���ȃG���[���������܂���" & vbCrLf & _
            "�u�Ǎ��ݒ�v�V�[�g�̓��e���m�F���Ă�������", _
            vbCritical, _
            "�u�Ǎ��ݒ�v�V�[�g�@�ǂݎ��@�G���["
    
End Function


'====================================================================
'�@�t�@�C���̃p�X���擾����֐�
'�@�Ԃ�l�FTrue/False (Boolean�^)
'�@�@�@�@�@True �F�p�X���擾�ł����ꍇ
'�@�@�@�@�@False�F�p�X���擾�ł��Ȃ������ꍇ
'�@�����@�FFile_Path�@�t�@�C���̃p�X
'�@�@�@�@�@�@�@�@�@�@(ByRef�w��ɂ��)�擾�����p�X���i�[���ĕԂ�
'�@�ŗL�ϐ��FfFlg�@���s�ς݃t���O
'====================================================================

Private Function FnFilePicker(ByRef File_Path As String) As Boolean
    
    Static fFlg As Boolean
    
    With Application.FileDialog(msoFileDialogFilePicker)
        
        '����̂݁A���̃t�@�C���̃p�X�������t�H���_�Ƃ���
        If Not fFlg Then
            .InitialFileName = ThisWorkbook.Path
            fFlg = True
        End If
        
        With .Filters   '�I���\�t�@�C���̍i�荞��
            .Clear
            .Add "CSV�t�@�C��", "*.csv,*.txt", 1
            .Add "�S�Ẵt�@�C��", "*.*", 2
        End With
        
        '�_�C�A���O���p�X���擾�B���ی��ʂ�ϐ��̌��ʂƂ��ĕԂ�
        If .Show = True Then
            File_Path = .SelectedItems(1)
            FnFilePicker = True
        End If
    End With
    
End Function

'====================================================================
'�@CSV�V�[�g�̏�����
'�@�i�K�v�ɉ����ē��e�ǉ��j
'====================================================================

Private Sub CSV_Initialaize()
    Worksheets("CSV").Cells.Delete         '�V�[�g�̒��g��S�č폜
End Sub



'====================================================================
'�@CSV��ǂݍ��ޏ���������v���V�[�W���i�g���܂킵���������ߓƗ�)
'�@�����@�FFile_Path�@CSV�t�@�C���̃p�X
'          CSV_Array�@CSV�Ǎ��̔z��
'�@�ϐ��@�FfFlg�@���s�ς݃t���O
'====================================================================

Private Sub CSV_read_Macro(File_Path As String, CSV_Array)
    Dim sh As Worksheet
    Set sh = Worksheets("CSV")
    
    With sh.QueryTables.Add( _
        Connection:="TEXT;" & File_Path, _
        Destination:=sh.Range("A1"))            'Connection:�ǂݍ��݃t�@�C���ADestination:�\�t����
        
        .Name = "temp"                          '����̓ǂݍ��ݑ���̖��́i�Ō�ɍ폜����̂łȂ�ł��悢�j
        .AdjustColumnWidth = True               '�񕝂̎����ݒ�
        .TextFilePlatform = 932                 '�����R�[�h�F932 SJIS
        .TextFileCommaDelimiter = True          '�J���}��؂�
        .TextFileColumnDataTypes = CSV_Array    '1=�����A2=������
        .Refresh BackgroundQuery:=False         '�o�b�N�O���E���h����(False:���Ȃ��B�o�b�N�O���E���h��������Ɠǂݍ��ݑO�Ƀ}�N�������֐i�ނ���)
        
        .Delete                                 '�N�G���̍폜(�f�[�^�ڑ�������)
    End With
    
    Set sh = Nothing
End Sub


'====================================================================
'�@�}�N�����s���ɂ悭�g���ݒ���܂Ƃ߂ď���
'�@�����@�FFlag�@True/False (Boolean�^)
'====================================================================

Property Let MacroMode(ByVal Flag As Boolean)
    With Application
        .EnableEvents = Not Flag            '�C�x���g�̎��s�E��~
        .ScreenUpdating = Not Flag          '��ʍX�V�̎��s�E��~
        .Calculation = IIf(Flag, xlCalculationManual, xlCalculationAutomatic)   '�Čv�Z�̎��s�E��~
    End With
End Property

