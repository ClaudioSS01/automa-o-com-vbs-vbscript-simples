'    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile("kwl.vbs").readAll()




'linha acima ja necessaria para integrar o kwl no documento vbs
'===========================================================================================================================================
'framework criado por claudio versao do projeto 0.0.1
'data de inicio da criação 11/05/2020
'atualização 14/05/2020

'=============================================================================================================================================
'exibe comandos de ajuda
dim pulaLinha : pulaLinha = ""&chr(13)&""
dim comandosAjuda : comandosAjuda = "PARA AJUDA: "&chr(13)&" help(), comandos(), comando() "&chr(13)&""&chr(13)&""

function ajuda()
WScript.Echo(""&novosComandos&""&comandosAjuda&""&comandosMatematica&""&comandosTexto&""&comandosAdicionarVbs&""&comandosAbrirProgramas&""&comandosApertarTeclas&""&comandosParaPararOvbs&""&comandosAtalhos&""&comandosDeTeclasEspeciais&""&comandoDeSetas&""&comandosChrome&"")
end function

function help()
ajuda()
end function

function comandos()
ajuda()
end function


dim novosComandos : novosComandos = "Para Notificar o usuario: "&chr(13)&" notificar('titulo da notificacao','texto da notificacao') "&chr(13)&""&chr(13)&""
novosComandos = novosComandos&"Para fazer ele falar useu o comando:  "&chr(13)&" fale('texto que eu quero que seja dito')"&chr(13)&""&chr(13)&""






function comando()
ajuda()
end function
'==========================================================================================================================================
'comandos do whatsapp
'https://api.whatsapp.com/send?phone=5511994953116&text=confirmo%20meu%20plantao%20dia%20taltaltal%20no%20hospital%20taltaltal%20codigodeprotocolo%20qazwsx

dim comandosWhatsapp : comandosWhatsapp = "COMANDOS WHATSAPP: "&pulaLinha&" enviarMensagem(numero,mensagem), OBS: numero nao pode começar com zero nem com 55 ex: 11 - 99495-3116"&pulaLinha&""&pulaLinha&""

function enviarMensagem(numero, mensagem)
dim mensagemTratado : mensagemTratado = Replace(texto," ","%20")
Dim oShell
Set oShell = WScript.CreateObject("WSCript.shell")
oShell.run "cmd /K start chrome --incognito https://api.whatsapp.com/send?phone=55"&numero&"&text="&mensagemTratado&" & exit", 0, false
end function


'==========================================================================================================================================
'comandos do chrome
'start chrome --incognito "http://www.iot.qa/2018/02/narrowband-iot.html"
dim comandosChrome : comandosChrome = "COMANDOS CHROME: "&pulaLinha&" chromeIrPara(url), fechaChrome(), fecharChrome(), "&pulaLinha&""&pulaLinha&""

function chromeIrPara(url)
dim urlTratado : urlTratado = Replace(texto,"https://","")
urlTratado : urlTratado = Replace(texto,"http://","")
Dim oShell
Set oShell = WScript.CreateObject("WSCript.shell")
oShell.run "cmd /K start chrome --incognito http://"&url&" & exit", 0, false
end function




'======================================================================
function cmd(comando)
dim comandoTratado : comandoTratado = Replace(comando,"cmd /K","")

Dim oShell
Set oShell = WScript.CreateObject("WSCript.shell")
oShell.run "cmd /K start "&comandoTratado&" & exit", 0, false
end function


function fechaChrome()
fecharPrograma("chrome")
end function

function fecharChrome()
fecharPrograma("chrome")
end function

'===========================================================================================================================================
'mostra resultado de conta
dim comandosMatematica : comandosMatematica = " RESULTADO MATEMATICA: "&chr(13)&" resultado 2+2 "&chr(13)&""&chr(13)&""

function resultado(numero)
msgbox numero
end function
'===========================================================================================================================================
'mostra texto
dim comandosTexto : comandosTexto = "MOSTRAR TEXTO:"&chr(13)&"  mostre(texto), exibir(texto), mensagem(text0), popup(text0) "&chr(13)&""&chr(13)&""

function mostre(texto)
dim textoTrabalhado : textoTrabalhado = Replace(texto,"pula linha",""&chr(13)&"")
textoTrabalhado = Replace(texto,"pule linha",""&chr(13)&"")
textoTrabalhado = Replace(texto,"pular linha",""&chr(13)&"")
msgbox(textoTrabalhado)
end function
function mostrar(texto)
mostre(texto)
end function
function exibir(texto)
mostre(texto)
end function
function mensagem(texto)
mostre(texto)
end function
function popup(texto)
mostre(texto)
end function
'=============================================================================================================================================
'inclui arquivos
dim comandosAdicionarVbs : comandosAdicionarVbs = "  ADICIONAR VBS OU BAT: "&chr(13)&" include(caminho do arquivo), incluirArquivo(caminho do arquivo), executeArquivo(caminho do arquivo)"&chr(13)&""&chr(13)&""

Sub incluirArquivo(fSpec)
    executeGlobal CreateObject("Scripting.FileSystemObject").openTextFile(fSpec).readAll()
End Sub

Sub include(fSpec)
  incluirArquivo(fSpec)
End Sub

Sub executeArquivo(fSpec)
    incluirArquivo(fSpec)
End Sub
'===============================================================================================================================================
'espera o tempo definido
dim comandosEsperar : comandosEsperar = "COMANDOS ESPERAR: "&chr(13)&" esperar(tempo), espere(tempo), aguardar(tempo), aguarde(espere)"&chr(13)&""&chr(13)&""

function espere(tempo)
if tempo = "" then
msgbox "função espere precisa de parametro de tempo exemplo: espere(1000) correspondente a 1 segundo"
else
set WshShell=WScript.CreateObject("WScript.Shell")
WScript.sleep tempo
 end if
 end function

function aguarde(tempo)
espere(tempo)
 end function

 function aguardar(tempo)
espere(tempo)
 end function

 function esperar(tempo)
espere(tempo)
 end function
'==================================================================================================================================================
'comandos para abrir programas
dim comandosAbrirProgramas : comandosAbrirProgramas = " PARA ABRIR PROGRAMAS: "&chr(13)&" abraPrograma(nomedoPrograma), abrirPrograma(nomeDoPrograma), executarPrograma(nomeDoPrograma), executePrograma(nomeDoPrograma), iniciarPrograma(nomeDoPrograma)"&chr(13)&""&chr(13)&""


function abraPrograma(nomeDoPrograma)
Dim oShell
Set oShell = WScript.CreateObject("WSCript.shell")
oShell.run "cmd /K start "&nomeDoPrograma&".exe & exit", 0, false
end function

function abrirPrograma(nomeDoPrograma)
abraPrograma(nomeDoPrograma)
end function

function executarPrograma(nomeDoPrograma)
abraPrograma(nomeDoPrograma)
end function

function executePrograma(nomeDoPrograma)
abraPrograma(nomeDoPrograma)
end function

function iniciarPrograma(nomeDoPrograma)
abraPrograma(nomeDoPrograma)
end function

'===============================================================================================================================================
'comandos para apertar teclas simulando interface humana
dim comandosApertarTeclas : comandosApertarTeclas = " SIMULAR APERTAR TECLAS: "&chr(13)&"escrever(texto), escreva(texto), pressione(texto), pressionar(texto), envieTeclas(texto), "&chr(13)&""&chr(13)&""

function envieTeclas(texto)
'voc{^}e n{~}ao "&chr(180)&"e dr{(}a{)}
dim textoTratado : textoTratado = Replace(texto,"ê","^e")
textoTratado = Replace(texto,"ã","~a")
textoTratado = Replace(texto,"é","´e")
textoTratado = Replace(texto,"(","{(}")
textoTratado = Replace(texto,")","{)}")
set WshShell=WScript.CreateObject("WScript.Shell")
WshShell.sendkeys ""&textoTratado&""
end function

function escrever(texto)
envieTeclas(texto)
end function

function escreva(texto)
envieTeclas(texto)
end function

function pressione(texto)
envieTeclas(texto)
end function

function pressionar(texto)
envieTeclas(texto)
end function
'=================================================================================================================================================
'comandos para abrir uma caixa para parar o vbs
dim comandosParaPararOvbs : comandosParaPararOvbs = " PARA PARAR DE SCRIPT: "&chr(13)&" pararScript() "

function paravbs()
result = MsgBox ("aperte ok para parar os vbs", vbOkonly, "APERTE OK PARA PARAR O SCRIPT")
Select Case result
Case vbOk
Dim oShell
Set oShell = WScript.CreateObject ("WSCript.shell")
oShell.run "cmd /K taskkill /F /IM wscript.exe & taskkill /f /im mshta.exe & exit", 0, false
msgbox("FIM do SCRIPt")
End Select
end function


function pararScript()
paravbs()
end function

function paraScript()
paravbs()
end function

function pararVBS()
paravbs()
end function

'===========================================================================================================================================================
'comandos para fechar programas
dim comandosPararProgramas : comandosPararProgramas = "COMANDOS PARAR PROGRAMAS: "&chr(13)&" pararPrograma(nomeDoPrograma), paraPrograma(nomeDoPrograma), fechaPrograma(nomeDoPrograma), fecharPrograma(nomeDoPrograma)"

function fecharPrograma(nomeDoPrograma)
Dim oShell
Set oShell = WScript.CreateObject ("WSCript.shell")
oShell.run "cmd /K taskkill /F /IM "&nomeDoPrograma&".exe", 0, false
end function

function fechaPrograma(nomeDoPrograma)
fecharPrograma(nomeDoPrograma)
end function

function paraPrograma(nomeDoPrograma)
fecharPrograma(nomeDoPrograma)
end function

function pararPrograma(nomeDoPrograma)
fecharPrograma(nomeDoPrograma)
end function
'===========================================================================================================================================================
'comandos de atalhos
'  + igual shift
'  ^ igual control
'  % igual alt
' "" igual espaço
' colocar % {%}
dim comandosAtalhos : comandosAtalhos = "COMANDOS DE ATALHOS: "&chr(13)&"windowsR(), windowsD(), control(letra), controlShift(letra), shift(letra), alt(letra), controlAlt(letra), enter(), f1(), f12, del(), delete(),"&chr(13)&""&chr(13)&""

function windowsD()
     set objShell = CreateObject("shell.application")
          objShell.ToggleDesktop
       set objShell = nothing
end function

function control(letra)
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "^{letra}"
end function

function controlShift(letra)
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "+^{letra}"
end function

function controlAlt(letra)
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "%^{letra}"
end function

function alt(letra)
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "%{letra}"
end function

function altShift(letra)
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "%+{letra}"
end function

function windowsR()
Dim winR : winR = WScript.CreateObject ("WSCript.shell")
', 0, false oculta o cmd
windowsR.run "cmd /K explorer.exe Shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}", 0, false
end function

'===========================================================================================================================================================================================================
' comando de teclas especiais
dim comandosDeTeclasEspeciais : comandosDeTeclasEspeciais = "COMANDOS DE TECLAS ESPECIAIS: "&pulaLinha&" print(), printScreen(), enter(), delete(), del(), tab(), espaco()"&pulaLinha&""&pulaLinha&""

function print()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{PRTSC}"
end function

function espaco()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys " "
end function

function printscreen()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{PRTSC}"
end function

function enter()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{enter}"
end function

function detele()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{delete}"
end function

function del()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{del}"
end function

function tab()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{tab}"
end function

'============================================================================================================================================================================================================
'comandos de teclas de funções
dim comandoDeTeclasDeFuncoes : comandoDeTeclasDeFuncoes = "COMANDO DE TECLAS DE FUNCOES: "&pulaLinha&" F1(), F2() ATE F12() "&pulaLinha&""&pulaLinha&""

function f1()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F1}"
end function

function f2()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F2}"
end function

function f3()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F3}"
end function

function f4()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F4}"
end function

function f5()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F5}"
end function

function f6()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F6}"
end function

function f7()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F7}"
end function

function f8()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F8}"
end function

function f9()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F9}"
end function

function f10()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F10}"
end function

function f11()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F11}"
end function

function f12()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{F12}"
end function
'=========================================================================================================================================================================================
 'COMANDO DE SETAS

 dim comandoDeSetas : comandoDeSetas = "COMANDO DE SETAS: "&pulaLinha&" setaParaCima(), setaParaBaixo(), setaParaDireita(), setaParaEsquerda()"&pulaLinha&""&pulaLinha&""
 function setaParaCima()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{UP}"
 end function

  function setaParaBaixo()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{DOWN}"
 end function

  function setaParaDireita()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{RIGHT}"
 end function

  function setaParaEsquerda()
set WshShell=WScript.CreateObject("WScript.Shell")
espere(250)
WshShell.sendkeys "{LEFT}"
 end function

 '================================================================================================================================================================================

