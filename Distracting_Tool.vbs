Set wshshell = wscript.CreateObject("WScript.Shell") 'create object 
strText = "Congratulations, You've been hacked" 'text that needed to be played
Set objVoice = CreateObject("SAPI.SpVoice") 'create object of SAPI so that using this we can play the message
do 'loop starts from here
Wshshell.run "Notepad" 'open the notepad

wscript.sleep 10 'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "C" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "K" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "E" 'print the message on the screen
wscript.sleep 10 ' wait for 10 ms
wshshell.sendkeys "D" 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen

wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "H" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys "A" 'print the message on the screen
wscript.sleep 10'wait for 10 ms
wshshell.sendkeys " " 'print the message on the screen


objVoice.Speak strText ' play the text

'To blick the light of keyboard
wscript.sleep 5'wait for 5 ms
wshshell.sendkeys "{CAPSLOCK}"  'press the capslock 
wscript.sleep 5 'wait for 5 ms
wshshell.sendkeys "{NUMLOCK}" 'press the Numlock 
wscript.sleep 5'wait for 5 ms
wshshell.sendkeys "{SCROLLLOCK}" 'Press the scrolllock

loop 'loop ends here
