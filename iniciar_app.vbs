Dim sh
Set sh = CreateObject("WScript.Shell")

' Inicia o servidor Streamlit sem janela de console
sh.Run "python3.12 -m streamlit run " & Chr(34) & "C:\Users\ari\OneDrive\Documentos\SISTEMA ARI\GeraLaudo\main.py" & Chr(34) & " --server.headless true --browser.gatherUsageStats false", 0, False

' Aguarda o servidor subir e abre o browser
WScript.Sleep 4000
sh.Run "cmd /c start http://localhost:8501", 0, False
