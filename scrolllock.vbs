' ENCODING EUC-KR
Set ie = CreateObject("InternetExplorer.Application")
Set ws = CreateObject("WScript.shell")

' ���� ǥ��
' version info
version = "1.9"
copyright = "Copyleft �� 2023"

' ���̵� Ÿ��(��ũ�Ѷ� ��ư�� �ι� ���� �� �ٽ� ���������� ����ϴ� �ð�. �� ������ �ۼ�)
' idle for(a waiting-time after a set of scrolllock pressed. second)
idle = 570

' ������ �� (Switch �������� Ȯ���ϴ� ����. �и��� ������ �ۼ�)
' listener term (a listener checking if Switch is pressed or not. millisecond)
term = 1000

ie.Offline = True
ie.Navigate "about:blank"

' â�� ũ��
' size of window
ie.Height = 250
ie.Width = 300

' ���ʿ� ���������� ����
' unessential browser tools disabled
ie.Menubar = False
ie.StatusBar = False
ie.AddressBar = False
ie.Toolbar = False
ie.Visible = True

' ���� ���α׷� ���� ��ư�� Ŭ�� ���θ� üũ
' check if button Close clicked
Sub check_cleanup
	If ie.document.all("CloseVal").value = "false" Then
		cleanup
	End If
End Sub

' ���α׷��� ����
' terminate program
Sub cleanup
	ie.Quit
	WScript.Quit
End Sub

' ȭ�� �ε尡 ���������� 100�и��� ������ ���.
' wait until window completely loaded. keep wait for every 100 millisec
Do While ie.Busy : WScript.Sleep 100 : Loop

running = "<button name='Switch' onClick=document.all('SwitchVal').value='stop' style=""height:100px;width:150px"" id='Switch'>������</button>" _
	& "</br><button name='Close' onClick=document.all('CloseVal').value='false' style=""height:50px;width:150px"">�������α׷� ����</button>" _
	& "<input name='SwitchVal' value='start' type='hidden' id='SwitchVal' disabled>" _
	& "<input name='CloseVal' value='true' type='hidden' id='CloseVal' disabled>" _
	& "</br><input style=""font-family:'Malgun Gothic', Serif; width:170px"" id='Copyright' name='Copyright' value='Ver" _
	& version _
	& " Copyleft �� 2023' readonly disabled></style>"
	
ceased = "<button name='Switch' onClick = document.all('SwitchVal').value='start' style=""height:100px;width:150px"" id='Switch'>�ߴܵ�</button>" _
	& "</br><button name='Close' onClick = document.all('CloseVal').value='false' style=""height:50px;width:150px"">�������α׷� ����</button>" _
	& "<input name='SwitchVal' value='stop' type='hidden' id='SwitchVal' disabled>" _
	& "<input name='CloseVal' value='true' type='hidden' id='CloseVal' disabled>" _
	& "</br><input style=""font-family:'Malgun Gothic', Serif; width:170px"" id='Copyright' name='Copyright' value='Ver" _
	& version _
	& " Copyleft �� 2023' readonly disabled></style>"
	
' â �Ӽ�
' window title and inital status
ie.document.Title = "ȭ�� ȣȣ��"
ie.document.body.innerHTML = running

Do While true
	' ���� input ���� stop ���� �ٲ𶧱��� ����
	' keep this loop until SwitchVal is changed to stop
	Do While ie.document.all("SwitchVal").value = "start"
		'{SCROLLLOCK}�� ����� ��.
		'use scrolllock only.
		ws.SendKeys"{SCROLLLOCK}{SCROLLLOCK}"
		ie.document.body.innerHTML = running
		
		x = 0
		Do While x < idle
			On Error Resume Next
			' ��ư�� �������� ���� Ȯ��
			' check if Switch clicked
			If ie.document.all("SwitchVal").value = "stop" Then
				x = x + idle
				ie.document.body.innerHTML = ceased
				If Err.Number <> 0 Then
					ie.Quit
					WScript.Quit
					Exit Do
				End If
			End If
			
			' ������ �����ư�� �������� Ȯ��
			' check if Close clicked
			check_cleanup
			
			' �ȴ����ٸ� ������ �� ��ŭ ���
			' if not, wait for term
			WScript.Sleep term
			x = x + 1
		Loop
		
		' ������ �����ư�� �������� Ȯ��
		' check if Close clicked
		check_cleanup
	Loop
	
	' �ߴܵ� ���¿��� �����ư�� �������� Ȯ��
	' check if Close clicked
	check_cleanup
	WScript.Sleep term
Loop