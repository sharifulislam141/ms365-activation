## MS Office, Excel , Powerpoint etc Activation Instructions

### Step 1: Open Command Prompt in Administrator Mode
First, you need to open Command Prompt with admin rights, then follow the instructions below step by step. Just copy/paste the commands and do not forget to hit Enter to execute them.

### Step 2: Open the Location of the Office Installed on Your PC
```
cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16
```
###Step 3: Convert Your Retail License to a Volume One

```
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
```
###Step 4: Activate Your Office Using KMS Client Key

```
cscript ospp.vbs /setprt:1688
cscript ospp.vbs /unpkey:6MWKP >nul
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
cscript ospp.vbs /sethst:107.175.77.7
cscript ospp.vbs /act
```
If you see the error 0xC004F074, it means that your internet connection is unstable or the server is busy. Please make sure your device is online and try the command cscript ospp.vbs /act again until you succeed.

