#TODO: Get computer name from zip, open zip, parse through data,
# attach result (decrytion enabled, disabled, etc) to a single file.


import re
import win32com.client
from amp_output_pull import amp_output_pull

folder = 'C:\\temp\\test\\'

amp_output_pull(folder, 'Auto Policy', 'Processed')