import urllib
import subprocess
import time


ncentral_url = "https://advantage.purpleguys.com/download/2022.2.0.77/winnt/N-central/WindowsAgentSetup.exe"
local_agent = "C:\\temp\\WindowsAgentSetup.exe"
urllib.urlretrieve(ncentral_url, local_agent)

time.sleep(60)

f = open("C:\\temp\\agent_run.bat", "w+")
f.write(r'C:\temp\WindowsAgentSetup.exe /S /quiet /V" /qn CUSTOMERID=2007 CUSTOMERNAME=\"Solar Alternatives\" REGISTRATION_TOKEN=db308236-831b-5363-47b7-1026c2129f64 SERVERPROTOCOL=HTTPS SERVERADDRESS=advantage.purpleguys.com SERVERPORT=443 "')
f.close()

subprocess.run(["C:\\temp\\agent_run.bat"])