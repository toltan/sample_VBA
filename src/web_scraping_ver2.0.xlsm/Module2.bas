Attribute VB_Name = "Module2"
Public Sub DataScraping() 'web�y�[�W��������擾 2016/08/11
Dim nextro As Integer

Dim objE As InternetExplorer
Dim htmlDoc As HTMLDocument, htmlDoc2 As HTMLDocument, htmlDoc3 As HTMLDocument
Dim el As IHTMLElement, fl As IHTMLElement, fl2 As IHTMLElement
Dim colTR, colTH, colTD, h2tag As IHTMLElementCollection
Dim col2L, colScript, colRotation, colNoNo As IHTMLElementCollection
Dim urls As New Collection
Dim urcol As Variant, counter As Variant, varCounter As Variant
Dim urlNo As Integer, nextUrl As Integer
Dim a, b, c, col, ro As Integer

Set objE = CreateObject("Internetexplorer.Application") 'IE�I�u�W�F�N�g�̐���
nextro = 2
urlNo = 4
nextUrl = 0

Worksheets("scraiping").Range(Cells(1, 2), Cells(Rows.Count, Columns.Count)).ClearContents
Application.ScreenUpdating = False '������\���ɂ���
objE.navigate "http://********************************" '�T�C�gURL

Do While objE.Busy = True Or objE.readyState < READYSTATE_COMPLETE '�X�V�\�ɂȂ�܂�DoEvents���J��Ԃ�
    DoEvents
Loop

Set htmlDoc = objE.document

For Each el In htmlDoc.Links '�y�[�W��HTML��a�^�O�����W
    urls.Add (el.href) '�ϐ�urls�Ƀ����NURL�������Ă���
Next el

For Each urcol In urls
    
    If objE.LocationURL = "http://daidata.goraggio.com/100185/list/?type=3&bt=21.30&f=2#HeaderWrapper&hist_num=1" Then
        objE.Visible = True '�X�N���C�s���O���I�������E�B���h�E��\�����A�������b�Z�[�W��\��
        MsgBox "scraiping complete!", vbInformation
        Exit Sub
    End If
        
    objE.navigate urls(urlNo + nextUrl) & "&hist_num=1"
    
    Do While objE.Busy = True Or objE.readyState < READYSTATE_COMPLETE '�҂�
        DoEvents
    Loop
    '�����y�[�W���łȂ������珈�������Ȃ��Ŏ��̃y�[�W��
    If Not objE.LocationURL Like "http://daidata.goraggio.com/100185/unit_list/?model*" Then
        GoTo err
    End If
        
    Do While objE.Busy = True Or objE.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    Set htmlDoc2 = objE.document
    Set colTR = htmlDoc2.getElementsByTagName("tr") '���݂̃y�[�W��tr�^�O���擾
    Set colTH = htmlDoc2.getElementsByTagName("th") '���݂̃y�[�W��th�^�O���擾
    Set h2tag = htmlDoc2.getElementsByTagName("strong") '���݂̃y�[�W��strong�^�O���擾
    
    col = 0
    Worksheets("scraiping").Cells(nextro - 1, 2).Value = h2tag(0).innerText
        
        For Each fl In colTR 'tr���̃f�[�^���擾
            Set colTD = fl.getElementsByTagName("td") 'tr����td�����o���B
            On Error Resume Next
            a = 0
            col = col + 1
            ro = nextro
            
            For c = 0 To 4 '���tr��td��9���邽��(���m�ɂ�0�Ԗڂɋ��td������)
            
                Select Case c '1,2,3,4���V�[�g�ɓ]�L
                    Case c = 0, 1, 2, 3, 4 '
                    Worksheets("scraiping").Cells(ro, col).Value = colTD(a).innerText
                    ro = ro + 1 '�s��1���炷
                End Select
        
            a = a + 1 '����td�v�f
            Next c

        Next fl
        
    objE.navigate urls(urlNo + nextUrl) & "&hist_num=1&disp=2&graph=1"
    
    Do While objE.Busy = True Or objE.readyState < READYSTATE_COMPLETE '�҂�
        DoEvents
    Loop
    
    Set htmlDoc3 = objE.document
    Set col2L = htmlDoc3.getElementById("Main-Contents")
    Set colScript = col2L.getElementsByTagName("script")
    Set colRotation = htmlDoc3.getElementsByClassName("Text-Green today")
    Set colNoNo = htmlDoc3.getElementsByClassName("Radius-Slot")
    col = 1
    a = 0
    
        For Each fl2 In colRotation
            col = col + 1
            Worksheets("scraiping").Cells(ro, col).Value = colRotation(a).innerHTML
            a = a + 1
        Next
        
    ro = ro + 1 '�s��1���炷
    col = 1
    a = 0
        
    For Each fl2 In colRotation
        col = col + 1
        counter = Split(colScript(a).innerHTML, Chr(10))
        counter = Split(counter(5), "]")
        counter = Split(counter(UBound(counter) - 2), ",")
        Worksheets("scraiping").Cells(ro, col).Value = counter(UBound(counter))
        a = a + 1
    Next
    
    nextro = ro + 2
    
err:
    nextUrl = nextUrl + 1
    Application.Wait Now + TimeValue("00:00:05")
    DoEvents '�҂�
    
Next urcol

Application.ScreenUpdating = True '�����\������

End Sub



