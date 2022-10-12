from urllib import request
import subprocess
import time


ncentral_url = "https://advantage.purpleguys.com/download/2022.2.0.77/winnt/N-central/WindowsAgentSetup.exe"
local_agent = "C:\\temp\\WindowsAgentSetup.exe"
request.urlretrieve(ncentral_url, local_agent)

time.sleep(60)

subprocess.run(["C:\\temp\\WindowsAgentSetup.exe", "/quiet",
               "/v",  "/qn", "CUSTOMERID=2007",
                "CUSTOMERNAME=Solar Alternatives",
                "REGISTRATION_TOKEN=db308236-831b-5363-47b7-1026c2129f64",
                "SERVERPROTOCOL=HTTPS",
                "SERVERADDRESS=advantage.purpleguys.com",
                "SERVERPORT=443"])