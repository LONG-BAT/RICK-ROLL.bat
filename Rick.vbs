While true
Dim oPlayer
set oPlayer = CreateObject("WMPlayer.OCX")
oPlayer.URL = "https://dl.dropboxusercontent.com/s/wza1sr5ot55bzdb/rickroll.mp3?dl=0"
oPlayer.controls.play
While oPlayer.playState <> 1 ' 1 = Stopped
WScript.Sleep 100
wend
oplayer.close
wend
