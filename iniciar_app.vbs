Dim sh
Set sh = CreateObject("WScript.Shell")

' cmd /c resolve o python3.12 do Windows Store corretamente
sh.Run "cmd /c python3.12 -m streamlit run " & Chr(34) & "C:\Users\ari\OneDrive\Documentos\SISTEMA ARI\GeraLaudo\main.py" & Chr(34) & " --server.headless true --browser.gatherUsageStats false", 0, False

' Aguarda o servidor subir e abre o browser
WScript.Sleep 5000
sh.Run "cmd /c start http://localhost:8501", 0, False
