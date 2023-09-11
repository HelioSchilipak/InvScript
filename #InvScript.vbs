On Error Resume Next


'------------------------------- Desativar Windows Defender

Wscript.Echo ("|||     Seja bem vindo ao InvScript© 2020 - 2023     |||" & vbCrLf & vbCrLf & "- Desative a PROTEÇÃO EM TEMPO REAL do Windows Defender")


'------------------------------- Verificar arquivo PK
Dim verificarArquivo

	Set verificarArquivo = CreateObject("Scripting.FileSystemObject") 

	if verificarArquivo.FileExists(".\Model\pk\pk.exe") = false Then
        Msgbox ("ATENÇÃO!!!" & vbCrLf  & vbCrLf &"- Arquivo pk.exe não foi localizado;" & vbCrLf  & vbCrLf & "- Será realizado uma nova Extração de pk.zip; " & vbCrLf & vbCrLf & "- Selecione a opção Substituir todos os arquivos no destino. ")
    

' Exemplo de script VBScript para extrair um arquivo ZIP

Dim zipFilePath
zipFilePath = "\\srvcoat\SERVIDOR\Inventario\Helio\Model\pk.zip"

Dim destinationFolder
destinationFolder = "\\srvcoat\SERVIDOR\Inventario\Helio\Model\"

Set objShell = CreateObject("Shell.Application")
Set zipFile = objShell.NameSpace(zipFilePath)
Set destination = objShell.NameSpace(destinationFolder)

destination.CopyHere zipFile.Items

Set objShell = Nothing
Set zipFile = Nothing
Set destination = Nothing

MsgBox "Arquivo ZIP extraído com sucesso!"

End if


'------------------------------- Inserção do Nome do Técnico:
Dim tecname
tecname=inputbox ("Insira o nome do Técnico Responsável:","ICI - COAT-ATH | InvScript© 2020 - 2023")
IF IsEmpty(tecname) Then
Msgbox ("Processo Cancelado.")
WScript.Quit
End If
tecname=UCase(tecname)
'------------------------------- Inserção do Nome da Secretaria

Dim secname
secname=inputbox ("Insira a secretaria a qual a máquina pertence:","ICI - COAT-ATH | InvScript© 2020 - 2023")
IF IsEmpty(secname) Then
Msgbox ("Processo Cancelado.")
WScript.Quit
End If
secname=UCase(secname)

'------------------------------- Inserção Número do Incidente / Requisição

Dim increqname
increqname=inputbox ("Insira o Número do Incidente / Requisição:","ICI - COAT-ATH | InvScript© 2020 - 2023")
IF IsEmpty(increqname) Then
Msgbox ("Processo Cancelado.")
WScript.Quit
End If
increqname=UCase(increqname)

'------------------------------- Inserção Número do Inventário

Dim invname
invname=inputbox ("Insira o Número do Inventário:","ICI - COAT-ATH | InvScript© 2020 - 2023")
IF IsEmpty(invname) Then
Msgbox ("Processo Cancelado.")
WScript.Quit
End If
invname=UCase(invname)


'------------------------------- Conferência dos dados inseridos


WScript.Echo "Confira os dados inseridos: "&chr(13)&chr(13) &"TÉCNICO:  "& tecname &chr(13)&"SECRETARIA:  "& secname &chr(13)&"INVENTÁRIO:  "& invname &chr(13)&"INC/REQ:  "& increqname &chr(13)



'----------------------------- Buscar Hostname
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
"SELECT * FROM Win32_ComputerSystem",,48)
For Each objItem in colItems
nomepc = objItem.Caption
Next



'----------------------------- Criar o arquivo

Dim fso, txtfile, nomearquivo
Set fso = CreateObject("Scripting.FileSystemObject")

nomearquivo =  increqname  & "-" &nomepc & " (" & Day(Now) & "." & Month(Now) & "." & Year(Now)& " " & Hour(Now)& "h" & Minute(Now) & "m)" & ".txt"

Set txtfile = fso.CreateTextFile(nomearquivo , True)



'Set txtfile = fso.CreateTextFile(".\" &increqname  & "-" &nomepc & " (" & Day(Now) & "." & Month(Now) & "." & Year(Now)& " " & Hour(Now)& "h" & Minute(Now) & "m)" & ".txt", True)

'----------------------------- Cabeçalho Documento e Impressão de Informações Inseridas 

txtfile.Write ("==========================================================")
txtfile.WriteBlankLines(1)
txtfile.Write ("|| GERÊNCIA DE INFRAESTRUTURA E SUPORTE TÉCNICO - GESUP ||")
txtfile.WriteBlankLines(1)
txtfile.Write ("|| COORDENAÇÃO DE ASSISTÊNCIA TÉCNICA - COAT            ||")
txtfile.WriteBlankLines(1)
txtfile.Write ("|| PLANILHA PADRÃO ALTERAÇÕES DE HARDWARE / SOFTWARE    ||")
txtfile.WriteBlankLines(1)
txtfile.Write ("==========================================================")
txtfile.WriteBlankLines(1)

txtfile.write ("IT2M-"&increqname  & "-" &nomepc & " (" & Day(Now) & "." & Month(Now) & "." & Year(Now)& " " & Hour(Now)& "h" & Minute(Now)  & "m)")

txtfile.WriteBlankLines(2)
txtfile.write ("|| REQ/INC___________________ "&increqname)
txtfile.WriteBlankLines(1)

txtfile.write ("|| INVENTÁRIO________________ "&invname)
txtfile.WriteBlankLines(1)

txtfile.write ("|| SECRETARIA________________ "&secname)
txtfile.WriteBlankLines(1)

txtfile.write ("|| HOSTNAME__________________ "&nomepc)
txtfile.WriteBlankLines(1)

txtfile.write ("|| TÉCNICO___________________ "&tecname)
txtfile.WriteBlankLines(1)  

txtfile.write ("|| REQ. ATUALIZAÇÃO__________ IT2M-")
txtfile.WriteBlankLines(1) 
txtfile.WriteBlankLines(1)




'--------------------------------------------Ler CHAVES 
'--------------------------------------------Decisão ler chave ou não
Dim WhShell, BttCode
Set WhShell = WScript.CreateObject("WScript.Shell")

BttCode = WhShell.Popup("Gostaria de capturar as licenças do Windows e Office?" & vbCrLf & vbCrLf & "Caso ocorra o bloqueio da captura das chaves pelo Anti-Vírus:"  & vbCrLf & vbCrLf &  "- Adicione a exceção para execução."& vbCrLf & "- Cancele as execuções em andamento."& vbCrLf & "- Reinicie o processo.", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)

Select Case BttCode
	
case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit
		
case 6   


'----------------------------- Inicio Impressão e Captura das Licenças

txtfile.Write ("==================================================")
txtfile.WriteBlankLines(2)
txtfile.Write("|OFFICE & WINDOWS|")
txtfile.WriteBlankLines(2)


	'Executar cmd chave office
	CreateObject("WScript.Shell").Popup "Capturando Licenças. Aguarde...", 3, "ICI - COAT-ATH | InvScript© 2020 - 2023"
	set objshell=WScript.CreateObject("Wscript.shell")
	ObjShell.run".\Model\pk\pk.exe /WindowsKeys 1 /OfficeKeys 1 /IEKeys 0 /SQLKeys 0 /ExchangeKeys 0 /ExtractEdition 0 /stext .\Model\chavetemp.txt"
	Set objshell=Nothing

	'Verificar se arquivo de chave foi criado
	'-----------------------------------------------------------
	Dim oFSO, oTxtFile

	Set oFSO = CreateObject("Scripting.FileSystemObject") 
	
	'Msgbox ("Criando arquivo da licença de office... Pressione OK")
	Do While oFSO.FileExists(".\Model\chavetemp.txt") = false
         
         Loop             
   	
	Set objShell = WScript.CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Const ForReading = 1
	Set objTxtServers = objFSO.OpenTextFile (".\Model\chavetemp.txt", ForReading)
	Do Until objTxtServers.AtEndOfStream
	strLine = objTxtServers.ReadLine
	If InStr(strLine,"Windows ") > 0 Then
	txtfile.WriteBlankLines(1)
	txtfile.Write ("WINDOWS -------------------------------------------")
	txtfile.WriteBlankLines(1)
	txtfile.Write(strLine)
	txtfile.WriteBlankLines(1)
	ElseIf InStr(strLine,"Product Name") > 0 Then
	txtfile.Write ("Office----------------------------------------------")
	txtfile.WriteBlankLines(1)
	txtfile.Write(strLine)
	txtfile.WriteBlankLines(1)
	'---------------------------------- Verificação Chave Windows Criptografada
	ElseIf InStr(strLine,"VK7JG-NPHTM") > 0 Then
	txtfile.Write(strLine)
	txtfile.Write("  --> Licença Criptografada.")
	txtfile.WriteBlankLines(1)
	txtfile.Write("Product Key       : --> OBRIGATÓRIO adicionar a chave utilizada para ativação")
	txtfile.WriteBlankLines(1)
	ElseIf InStr(strLine,"NF6HC-QH89W-F8WYV-WWXV4-WFG6P") > 0 Then
	txtfile.Write(strLine)
	txtfile.Write("  --> Licença Criptografada.")
	txtfile.WriteBlankLines(1)
	txtfile.Write("Product Key       : --> OBRIGATÓRIO adicionar a chave utilizada para ativação")
	txtfile.WriteBlankLines(1)
	'---------------------------------- Verificação Chave Windows Select
	ElseIf InStr(strLine,"Product key was not found") > 0 Then
	txtfile.Write(strLine)
	txtfile.Write("  --> Licença Criptografada.")
	txtfile.WriteBlankLines(1)
	txtfile.Write("Product Key       : --> OBRIGATÓRIO adicionar a chave utilizada para ativação")
	txtfile.WriteBlankLines(1)
	'---------------------------------- Verificação Chave Office Select
	ElseIf InStr(strLine,"YC7DK-G2NP3") > 0 Then
	txtfile.Write(strLine)
	txtfile.Write("  --> Licença Criptografada.")
	txtfile.WriteBlankLines(1)
	txtfile.Write("Product Key       : --> OBRIGATÓRIO adicionar a chave utilizada para ativação")
	txtfile.WriteBlankLines(1)
	'---------------------------------- Verificação Chave Office Select
	ElseIf InStr(strLine,"VYBBJ-TRJPB") > 0 Then
	txtfile.Write(strLine)
	txtfile.Write("  --> Licença Criptografada.")
	txtfile.WriteBlankLines(1)
	txtfile.Write("Product Key       : --> OBRIGATÓRIO adicionar a chave utilizada para ativação")
	txtfile.WriteBlankLines(1)
	ElseIf InStr(strLine,"Product Key") > 0 Then
	txtfile.Write(strLine)
	txtfile.WriteBlankLines(1)
	End if
	Loop
	objTxtServers.close
	txtfile.WriteBlankLines(1)
	txtfile.Write ("---------------------------------------------------")
	txtfile.WriteBlankLines(1)


	'----------------------------------------- Deletar arquivo temporário
	Set obj = CreateObject("Scripting.FileSystemObject") 
	Dim filename

	obj.DeleteFile ".\Model\chavetemp.txt"
	Set obj=Nothing

'----------------------------------------------------MAC


Case 7   
End Select
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(1)

'Descobrir sistema
strComputer = "."
strProperties = "*"'"CSName, Caption, OSType, Version, OSProductSuite, BuildNumber, ProductType, OSLanguage, CSDVersion, InstallDate, RegisteredUser, Organization, SerialNumber, WindowsDirectory, SystemDirectory"
objClass = "Win32_OperatingSystem"
strQuery = "SELECT " & strProperties & " FROM " & objClass
Set colOS = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)
For Each objItem in colOS
sistema = objItem.Caption
next 


'If Windows XP
if sistema = "Microsoft Windows XP Professional" then 
strQuery = "SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID > ''"
Set objWMIService = GetObject( "winmgmts://./root/CIMV2" )
Set colItems      = objWMIService.ExecQuery( strQuery, "WQL", 48 )
txtfile.write ("|MAC|")
contatodormac = 0
For Each objItem In colItems
contadormac = contadormac + 1
if not isnull(objItem.MACAddress) then txtfile.write (vbCrLf & "MAC " & contadormac & ": " & objItem.MACAddress)
Next
txtfile.WriteBlankLines(2)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(2)
Else 
    txtfile.WriteBlankLines(1)
    txtfile.write ("|MAC|")
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter where physicaladapter=true")
    for Each objItem in colItems
        if not isnull(objItem.MACAddress) then txtfile.write (vbCrLf & objItem.description & ": " & objItem.MACAddress)
        next 
txtfile.WriteBlankLines(2)
    txtfile.Write ("==================================================")
txtfile.WriteBlankLines(2)
    End If


'-------------------------------------------------------------------- Captura dados Placa Mãe
strComputer = "." 
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard") 
txtfile.write("|PLACA MÃE|")
txtfile.WriteBlankLines(1)
For Each objItem in colItems 
    placamae = objItem.Manufacturer
    modelo = objItem.Product
    txtfile.write(placamae &"-"& modelo)
Next
txtfile.WriteBlankLines(2)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(2)

'------------------------------- Captura dados Processador
txtfile.write ("|PROCESSADOR|")
txtfile.WriteBlankLines(1)
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
"SELECT * FROM Win32_Processor",,48)
For Each objItem in colItems


'------------------------------------------------- Impressão Nome do Processador
txtfile.write(objItem.name)
txtfile.WriteBlankLines(2)
Next
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(2)
'----------------------------------Captura e Impressão Memória
txtfile.write ("|MEMORIA|")
txtfile.WriteBlankLines(1)
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
"SELECT * FROM Win32_physicalmemory",,48)
'Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
cont = 0
memoriatotal = 0
For Each objItem in colItems

cont = (cont + 1)
txtfile.write ("Modulo " & cont & ": " & objItem.capacity/1048576 & " MB")
memoriatotal = (objItem.capacity/1048576 + memoriatotal) 
txtfile.WriteBlankLines(1)
Next
txtfile.write("Memoria total: " & (memoriatotal/1024) &" GB")
txtfile.WriteBlankLines(2)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(2)

'---------------------------------- Captura e Impressão HD / SSD

txtfile.write ("|HD/SSD| ")
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
"SELECT * FROM Win32_diskdrive",,48)
contadorhd = 0
For Each objItem in colItems
'------------------------------------------------- Modelo do disco
'txtfile.write ("Disco:")
'txtfile.WriteBlankLines(1)
'txtfile.write (objItem.caption)
'txtfile.WriteBlankLines(1)
'----------------------------------------------------- Interface
'txtfile.write ("Interface:")
'txtfile.WriteBlankLines(1)
'txtfile.write (objItem.interfacetype)
txtfile.WriteBlankLines(1)
contadorhd = (contadorhd + 1)
txtfile.write ("Disco "& contadorhd)

'----------------------------------------------------- Capacidade
capacidade = int(objItem.size/1073741824)
If capacidade > 900 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 1 TB")
ElseIf capacidade > 695 And capacidade < 750 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 750 GB")
ElseIf capacidade > 400 And capacidade < 500 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 500 GB")
ElseIf capacidade > 231 And capacidade < 250 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 250 GB")
ElseIf capacidade > 225 And capacidade < 240 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 240 GB")
ElseIf capacidade > 140 And capacidade < 160 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 160 GB")
ElseIf capacidade > 110 And capacidade < 120 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 120 GB")
ElseIf capacidade > 70 And capacidade < 80 Then
txtfile.WriteBlankLines(1)
txtfile.write ("Capacidade: 80 GB")
End If
txtfile.WriteBlankLines(1)
txtfile.write ("Tamanho Real: ")
txtfile.write (Int(objItem.size/1073741824) & " GB")
txtfile.WriteBlankLines(1)
txtfile.Write ("--------------------------------------------------")
Next
txtfile.WriteBlankLines(2)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(2)

'--------------------------------------------------------- Captura informações Placa de Vídeo
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_VideoController")

For Each objItem in colItems

txtfile.write ("Adaptador de Vídeo: " & objItem.Description)
Next
txtfile.WriteBlankLines(2)
txtfile.Write ("==================================================")
txtfile.WriteBlankLines(2)

'------------------------------------------------- Captura Nome do adaptador Rede

txtfile.write ("|IP|")
txtfile.WriteBlankLines(1)
strComputer = "."
strProperties = "Description, MACAddress, IPAddress, IPSubnet, DefaultIPGateway, DNSServerSearchOrder, DNSDomain, DNSDomainSuffixSearchOrder, DHCPEnabled, DHCPServer, WINSPrimaryServer, WINSSecondaryServer, ServiceName"
objClass = "Win32_NetworkAdapterConfiguration"
strQuery = "SELECT " & strProperties & " FROM " & objClass & " WHERE IPEnabled = True AND ServiceName <> 'AsyncMac' AND ServiceName <> 'VMnetx' AND ServiceName <> 'VMnetadapter' AND ServiceName <> 'Rasl2tp' AND ServiceName <> 'PptpMiniport' AND ServiceName <> 'Raspti' AND ServiceName <> 'NDISWan' AND ServiceName <> 'RasPppoe' AND ServiceName <> 'NdisIP' AND ServiceName <> ''"
Set colAdapters = objWMIService.ExecQuery(strQuery, , wbemFlagReturnImmediately + wbemFlagForwardOnly)

'------------------------------------------------- Impressão dados de rede

For Each objItem in colAdapters
'For Each objItem in colItems
'txtfile.write ("Adaptador:")
'txtfile.WriteBlankLines(1)
'txtfile.write (objItem.Description)
'txtfile.WriteBlankLines(1)

'------------------------------------------------- Captura e Impressão IP

'txtfile.write ("IP: ")
'txtfile.WriteBlankLines(1)
IP_Address = objItem.IPAddress
txtfile.write (IP_Address(i))
txtfile.WriteBlankLines(1)
Next

'Set objFSO = CreateObject("Scripting.FileSystemObject") 
'Set Shell = WScript.CreateObject("WScript.Shell") 
'Set ShellApplication = WScript.CreateObject("Shell.Application") 
'Set objNetwork = CreateObject("WScript.Network") 
 
'Set objFSO = CreateObject("Scripting.FileSystemObject") 
'Set Shell = WScript.CreateObject("WScript.Shell") 
'Set ShellApplication = WScript.CreateObject("Shell.Application") 
'Set objNetwork = CreateObject("WScript.Network") 
 
   Dim WshShell, BtnCode
Set WshShell = WScript.CreateObject("WScript.Shell")

'------------------------------------------------------------------ Inicio CHECKLIST


'------------------------------------------------------------------ Caso aperte SIM

BtnCode = WshShell.Popup("Gostaria de fazer o Check List?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)



Select Case BtnCode
	
case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

case 6      	
	
	
	'------------------------------------------------------------------ Lista de Softwares Específicos

txtfile.WriteBlankLines(1)
txtfile.WriteBlankLines(2)
txtfile.write ("==================================================")
txtfile.WriteBlankLines(2)
txtfile.write ("|Sistemas e Softwares Específicos|")
txtfile.WriteBlankLines(2)


If InStr(secname,"SMS") > 0 Then
		BtnCode = WshShell.Popup("Instalar BDE?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BDE") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select

		BtnCode = WshShell.Popup("Instalar FLASH 10?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Flash") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar BPA?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BPA") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar GTM?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("GTM") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar Meta4?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Meta4") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SGP?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SPG") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SCNES?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SCNES") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SinanNet?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SinanNet") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		

		BtnCode = WshShell.Popup("Instalar RASS?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("RASS") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		
'-----------------------------------------------------------------------SE Nome FAS   		
   		ElseIf InStr(secname,"FAS") > 0 Then
		BtnCode = WshShell.Popup("Instalar BDE?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BDE") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar GTM?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("GTM") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SGP?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SGP") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar Sistemas FAS?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Sistemas FAS") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
'---------------------------------------------------------------------------------------SE nome MASE
	ElseIf InStr(secname,"MASE") > 0 Then
		BtnCode = WshShell.Popup("Instalar BDE?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BDE") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar GTM?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("GTM") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SGP?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SGP") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar Atualizador?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Atualizador") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select

'-------------------------------------------------------------------------------------- SE nome SMMA
	ElseIf InStr(secname,"SMMA") > 0 Then
		BtnCode = WshShell.Popup("Instalar BDE?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BDE") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar GTM?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("GTM") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SGP?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SGP") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar Atualizador?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Atualizador") 
   					txtfile.WriteBlankLines(1)
   		case 7      
   		End Select
   		
   		BtnCode = WshShell.Popup("Instalar Localizador DLL?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Localizador DLL") 
   					txtfile.WriteBlankLines(1)
   		case 7      
   		
		End Select

'--------------------------------------------------------------------------------------- SE nome SMOP
	ElseIf InStr(secname,"SMOP") > 0 Then
		BtnCode = WshShell.Popup("Instalar BDE?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BDE") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar GTM?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("GTM") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SGP?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SGP") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar Atualizador?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Atualizador") 
   					txtfile.WriteBlankLines(1)
   		case 7      
   		End Select
   		
   		BtnCode = WshShell.Popup("Instalar Localizador DLL?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Localizador DLL") 
   					txtfile.WriteBlankLines(1)
   		case 7      
   		
		End Select
		
		BtnCode = WshShell.Popup("Instalar OCP DLL?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("OCP DLL") 
   					txtfile.WriteBlankLines(1)
   		case 7      
   		End Select
'------------------------------------------------------------------------------------- SE nome SMDT
	ElseIf InStr(secname,"SMDT") > 0 Then
		BtnCode = WshShell.Popup("Instalar BDE?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BDE") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar GTM?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("GTM") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SGP?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SGP") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar Coc Net?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Coc Net") 
   					txtfile.WriteBlankLines(1)
   		case 7      
   		End Select
   		
'-------------------------------------------------------------------------------------SE nome SMRH

	ElseIf InStr(secname,"SMRH") > 0 Then
		BtnCode = WshShell.Popup("Instalar BDE?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BDE") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar GTM?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("GTM") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SGP?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SGP") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar Business Object?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Business Object") 
   					txtfile.WriteBlankLines(1)
   		case 7      
   		End Select
   		
   		BtnCode = WshShell.Popup("Instalar Meta 4?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Meta 4") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
'---------------------------------------------------------------------------------------SE nome OUTROS
	Else 
		BtnCode = WshShell.Popup("Instalar BDE?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BDE") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar GTM?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("GTM") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		
		BtnCode = WshShell.Popup("Instalar SGP?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("SGP") 
   					txtfile.WriteBlankLines(1)
   		case 7      
		End Select
		


'---------------------------------------------------------------------------------------------
	
txtfile.WriteBlankLines(1)	
txtfile.write ("==================================================")	
txtfile.WriteBlankLines(2)	
txtfile.write ("|Check List|")
txtfile.WriteBlankLines(2)

		BtnCode = WshShell.Popup("Alterar Hostname (Secretaria-Inventário)?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Hostname:"+chr(9)+"Ok") 
   		case 7      txtfile.write ("Hostname:"+chr(9)+"--") 
		End Select
   		txtfile.write (chr(9))

		BtnCode = WshShell.Popup("Atualizar Drivers?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Atualização de Drivers: Ok") 
   		case 7      txtfile.write ("Atualização de Drivers: --") 
		End Select
		txtfile.write (chr(9))

		BtnCode = WshShell.Popup("Alocar Partição Unidade D:?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit
	
		case 6      txtfile.write ("Alocar Unidade D: Ok") 
   		case 7      txtfile.write ("Alocar Unidade D: --") 
		End Select
		txtfile.WriteBlankLines(1)

		BtnCode = WshShell.Popup("Atualizar BD Antivírus?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("BD Antivírus:"+chr(9)+"Ok") 
   		case 7      txtfile.write ("BD Antivírus:"+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))

		BtnCode = WshShell.Popup("Realizar Atualizações Automáticas?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("At. automáticas:"+chr(9)+"Ok") 
   		case 7      txtfile.write ("At. automáticas:"+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))

		BtnCode = WshShell.Popup("Ativar Windows?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Ativação Windows: Ok") 
   		case 7      txtfile.write ("Ativação Windows: --") 
		End Select
		txtfile.WriteBlankLines(1)

		BtnCode = WshShell.Popup("Ativar Office?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

		case 6      txtfile.write ("Atv. Office:"+chr(9)+"Ok") 
   		case 7      txtfile.write ("Atv. Office:"+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))

	
		
		BtnCode = WshShell.Popup("Alterar Papel de Parede?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Papel de parede:"+chr(9)+"Ok") 
   		case 7      txtfile.write ("Papel de parede:"+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))
		
		BtnCode = WshShell.Popup("Realizar Teste de Desempenho?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Desempenho:"+chr(9)+"  Ok") 
   		case 7      txtfile.write ("Desempenho:"+chr(9)+"  --") 
		End Select
		txtfile.WriteBlankLines(1)
		
		BtnCode = WshShell.Popup("Exibir ícone de Rede na Barra de tarefas?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

		case 6      txtfile.write ("Ícone Rede:"+chr(9)+"Ok") 
   		case 7      txtfile.write ("Ícone Rede:"+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))
		
		BtnCode = WshShell.Popup("Ajustar Hora Certa?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Hora certa:"+chr(9)+chr(9)+"Ok") 
   		case 7      txtfile.write ("Hora certa:"+chr(9)+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))
		
		BtnCode = WshShell.Popup("Testar Wi-Fi?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Testar wi-fi:"+chr(9)+"  Ok") 
   		case 7      txtfile.write ("Testar wi-fi:"+chr(9)+"  --") 
		End Select
		txtfile.WriteBlankLines(1)
		
		BtnCode = WshShell.Popup("Reconfigurar IP?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Reconfig. IP:"+chr(9)+"Ok") 
   		case 7      txtfile.write ("Reconfig. IP:"+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))
		
		BtnCode = WshShell.Popup("Remover Mídias?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Remover Mídias:"+chr(9)+chr(9)+"Ok") 
   		case 7      txtfile.write ("Remover Mídias:"+chr(9)+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))
		
		BtnCode = WshShell.Popup("Retornar Backup?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Retornar Backup:  Ok") 
   		case 7      txtfile.write ("Retornar Backup:  --") 
		End Select
		txtfile.WriteBlankLines(1)
		
		BtnCode = WshShell.Popup("Executar Samba Script?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Samba Script:"+chr(9)+"Ok") 
   		case 7      txtfile.write ("Samba Script:"+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))
		
		BtnCode = WshShell.Popup("Executar WSUS Script?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("WSUS Script:"+chr(9)+chr(9)+"Ok") 
   		case 7      txtfile.write ("WSUS Script:"+chr(9)+chr(9)+"--") 
		End Select
		txtfile.WriteBlankLines(1)
		
		BtnCode = WshShell.Popup("Fixar Lacre?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Fixar Lacre:"+chr(9)+"Ok") 
   		case 7      txtfile.write ("Fixar Lacre:"+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))

		BtnCode = WshShell.Popup("Será necessário criar Requisição de Atualização Inventário?", -1, "ICI - COAT-ATH | InvScript© 2020 - 2023", 3 + 32)
		Select Case BtnCode
		case 2		Msgbox ("Processo Cancelado.")

txtfile.close

'----------------------------------------- Deletar arquivo em caso de cancelamento
Set obj = CreateObject("Scripting.FileSystemObject") 

obj.DeleteFile nomearquivo
'obj.DeleteFile ".\teste.txt"
Set obj=Nothing

wscript.quit

   		case 6      txtfile.write ("Inventário:"+chr(9)+chr(9)+"Ok") 
   		case 7      txtfile.write ("Inventário:"+chr(9)+chr(9)+"--") 
		End Select
		txtfile.write (chr(9))

txtfile.WriteBlankLines(2)


	End If

txtfile.WriteBlankLines(1)

'----------------------------------------------------------Captura SOFTWARES Instalados	

txtfile.write ("==================================================")
txtfile.WriteBlankLines(2)
txtfile.write ("|SOFTWARES INSTALADOS|")
txtfile.WriteBlankLines(2)



CreateObject("WScript.Shell").Popup "Capturando Softwares Instalados. Aguarde...", 2, "ICI - COAT-ATH | InvScript© 2020 - 2023"
Const HKLM = &H80000002
Set objReg = GetObject("winmgmts://" & "." & "/root/default:StdRegProv")

writeList "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\", objReg, objFile
writeList "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\", objReg, objFile

Function writeList(strBaseKey, objReg, objFile) 
objReg.EnumKey HKLM, strBaseKey, arrSubKeys 
    For Each strSubKey In arrSubKeys
        intRet = objReg.GetStringValue(HKLM, strBaseKey & strSubKey, "DisplayName", strValue)
        If intRet <> 0 Then
            intRet = objReg.GetStringValue(HKLM, strBaseKey & strSubKey, "QuietDisplayName", strValue)
        End If
        objReg.GetStringValue HKLM, strBaseKey & strSubKey, "DisplayVersion", version
        objReg.GetStringValue HKLM, strBaseKey & strSubKey, "InstallDate", insDate 
        If (strValue <> "") and (intRet = 0) Then
            txtfile.write strValue & " - " & version & vbCrLf
        End If
    Next
End Function

 
'------------------------------------------------------------- Caso aperte não
'------------------------------------------------------------- Final da Impressão

Case 7   
End Select
txtfile.WriteBlankLines(1)
txtfile.WriteBlankLines(1)
txtfile.write("***********************Instituto das Cidades Inteligentes********************")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@          @@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@               @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@           @@@@@@@      @@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@         @@@@@@@@@@@@@@     @@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@         @@@@@@@@@@       @@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@             @@   @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@          @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@     @@@@  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@  @@@@@@@@  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ @@@@@    @@@@@  @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@   ICI - COAT-ATH | InvScript 2020 - 2023   @@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@   Desenvolvido por Daniel Bonato           @@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@   Colaboração Helio Schilipak              @@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)
txtfile.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
txtfile.WriteBlankLines(1)

Wscript.Echo "Informações adicionadas com Êxito!" & vbCrLf & vbCrLf & vbCrLf & "Desenvolvido por Daniel Bonato" & vbCrLf & "Colaboração Helio Schilipak" &vbCrLf & "ICI - COAT-ATH | InvScript© 2020 - 2023"
wscript.quit
