@echo off
cd /d "C:\Users\marco\OneDrive\Dokumente\ChatGPT\Bestellung\Solarabrechnung-fuer-Nutzer"
echo Solarabrechnung startet > "C:\Users\marco\OneDrive\Dokumente\ChatGPT\Bestellung\Solarabrechnung-fuer-Nutzer\solarabrechnung-start.log"
"C:\Program Files\WindowsApps\OpenAI.CodexPrimaryRuntime.v26-1105-1297-0_26.1105.1297.0_x64__3k8sg7r9htsxt\dependencies\python\python.exe" "C:\Users\marco\OneDrive\Dokumente\ChatGPT\Bestellung\Solarabrechnung-fuer-Nutzer\solar_server.py" >> "C:\Users\marco\OneDrive\Dokumente\ChatGPT\Bestellung\Solarabrechnung-fuer-Nutzer\solarabrechnung-start.log" 2>&1
