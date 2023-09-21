' ENCODING EUC-KR
Set ie = CreateObject("InternetExplorer.Application")
Set ws = CreateObject("WScript.shell")

' 버전 표시
' version info
version = "1.9"
copyright = "Copyleft ⓒ 2023"

' 아이들 타임(스크롤락 버튼을 두번 누른 뒤 다시 누를때까지 대기하는 시간. 초 단위로 작성)
' idle for(a waiting-time after a set of scrolllock pressed. second)
idle = 570

' 리스너 텀 (Switch 눌렀는지 확인하는 간격. 밀리초 단위로 작성)
' listener term (a listener checking if Switch is pressed or not. millisecond)
term = 1000

ie.Offline = True
ie.Navigate "about:blank"

' 창의 크기
' size of window
ie.Height = 250
ie.Width = 300

' 불필요 브라우저도구 제거
' unessential browser tools disabled
ie.Menubar = False
ie.StatusBar = False
ie.AddressBar = False
ie.Toolbar = False
ie.Visible = True

' 응용 프로그램 종료 버튼의 클릭 여부를 체크
' check if button Close clicked
Sub check_cleanup
	If ie.document.all("CloseVal").value = "false" Then
		cleanup
	End If
End Sub

' 프로그램을 종료
' terminate program
Sub cleanup
	ie.Quit
	WScript.Quit
End Sub

' 화면 로드가 끝날때까지 100밀리초 단위로 대기.
' wait until window completely loaded. keep wait for every 100 millisec
Do While ie.Busy : WScript.Sleep 100 : Loop

running = "<button name='Switch' onClick=document.all('SwitchVal').value='stop' style=""height:100px;width:150px"" id='Switch'>실행중</button>" _
	& "</br><button name='Close' onClick=document.all('CloseVal').value='false' style=""height:50px;width:150px"">응용프로그램 종료</button>" _
	& "<input name='SwitchVal' value='start' type='hidden' id='SwitchVal' disabled>" _
	& "<input name='CloseVal' value='true' type='hidden' id='CloseVal' disabled>" _
	& "</br><input style=""font-family:'Malgun Gothic', Serif; width:170px"" id='Copyright' name='Copyright' value='Ver" _
	& version _
	& " Copyleft ⓒ 2023' readonly disabled></style>"
	
ceased = "<button name='Switch' onClick = document.all('SwitchVal').value='start' style=""height:100px;width:150px"" id='Switch'>중단됨</button>" _
	& "</br><button name='Close' onClick = document.all('CloseVal').value='false' style=""height:50px;width:150px"">응용프로그램 종료</button>" _
	& "<input name='SwitchVal' value='stop' type='hidden' id='SwitchVal' disabled>" _
	& "<input name='CloseVal' value='true' type='hidden' id='CloseVal' disabled>" _
	& "</br><input style=""font-family:'Malgun Gothic', Serif; width:170px"" id='Copyright' name='Copyright' value='Ver" _
	& version _
	& " Copyleft ⓒ 2023' readonly disabled></style>"
	
' 창 속성
' window title and inital status
ie.document.Title = "화면 호호기"
ie.document.body.innerHTML = running

Do While true
	' 내부 input 값이 stop 으로 바뀔때까지 진행
	' keep this loop until SwitchVal is changed to stop
	Do While ie.document.all("SwitchVal").value = "start"
		'{SCROLLLOCK}만 사용할 것.
		'use scrolllock only.
		ws.SendKeys"{SCROLLLOCK}{SCROLLLOCK}"
		ie.document.body.innerHTML = running
		
		x = 0
		Do While x < idle
			On Error Resume Next
			' 버튼을 눌렀는지 먼저 확인
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
			
			' 실행중 종료버튼을 눌렀는지 확인
			' check if Close clicked
			check_cleanup
			
			' 안눌렀다면 리스너 텀 만큼 대기
			' if not, wait for term
			WScript.Sleep term
			x = x + 1
		Loop
		
		' 실행중 종료버튼을 눌렀는지 확인
		' check if Close clicked
		check_cleanup
	Loop
	
	' 중단됨 상태에서 종료버튼을 눌렀는지 확인
	' check if Close clicked
	check_cleanup
	WScript.Sleep term
Loop