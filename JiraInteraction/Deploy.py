# MSP Update deploy
# Creates deploy.zip and installs to MyDocuments\bin\MSP Update

import shutil
import os
import zipfile

SrcDrctry = "C:\\Users\\Bruce Pike Rice\\Documents\\Repos\\work-management\\"
Trgt = SrcDrctry + "Deploy\\"
InstllDrctry = "C:\\Users\\Bruce Pike Rice\\Documents\\bin\\MSP Update\\"
VmShrDrctry = "C:\\Users\\Bruce Pike Rice\\Documents\\Virtual Machines\\Shared With VMs\\"

# Copy release folder contents
RlsFldr = SrcDrctry + "JiraInteraction\\bin\\Release\\"
FlLst = os.listdir(RlsFldr)
for Itm in FlLst:
    #print ("item = %s" %(RlsFldr + Itm))
    shutil.copy(RlsFldr + Itm, Trgt)

# Copy other files
shutil.copy(SrcDrctry + "JiraInteraction\\UTS MSP Update Template.xlsm", Trgt)

shutil.copy(SrcDrctry + "JiraInteraction\\UTS MSP Update Readme.txt", Trgt)

# Create deploy.zip
shutil.make_archive(SrcDrctry + "deploy", 'zip', Trgt)

# Copy deploy zip to install directory and unzip
shutil.copy(SrcDrctry + "deploy.zip", InstllDrctry)
zip_ref = zipfile.ZipFile(InstllDrctry + "deploy.zip", 'r')
zip_ref.extractall(InstllDrctry + "Deploy\\")
zip_ref.close()

# Copy deploy zip to VM share directory
shutil.copy(SrcDrctry + "deploy.zip", VmShrDrctry)


wait = input("All done - PRESS ENTER TO CONTINUE.")