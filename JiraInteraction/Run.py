import subprocess
import os
VbxMngPth = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
VbxHdlssPth = "C:\Program Files\Oracle\VirtualBox\VBoxHeadless.exe"

subprocess.run([VbxMngPth, "startvm", "VM2"], shell=True)
#subprocess.run(["dir" , "/d", "/on"], shell=True)


#subprocess.run([VbxHdlssPth, "-s", "VM2"], shell=False)

#subprocess.run([VbxMngPth, "controlvm", "VM2", "savestate"], shell=False)

wait = input("All done - PRESS ENTER TO CONTINUE.")