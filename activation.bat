@echo off
setlocal EnableDelayedExpansion & cd /d "%~dp0"
title KMS激活脚本
%1 %2
ver|find "5.">nul&&goto :xptooff
mshta vbscript:createobject("shell.application").shellexecute("%~s0","goto :start","","runas",1)(window.close)&goto :eof

:start
cls
echo. & echo 正在检查激活服务器.....
echo.
ping kms.zbgw.xyz | find "超时"  > NUL &&  goto fail
ping kms.zbgw.xyz  | find "目标主机"  > NUL &&  goto fail



echo 已成功连接上服务器，开始检查系统激活情况，将自动跳过已永久激活的....
cscript //Nologo %windir%\system32\slmgr.vbs /xpr | find "已永久激活">NUL&&goto wintooff
ver | find "6.0." > NUL &&  goto winvista
ver | find "6.1." > NUL &&  goto win7
ver | find "6.2." > NUL &&  goto win8
ver | find "6.3." > NUL &&  goto win81
ver | find "10.0." > NUL &&  goto win10
echo 未找到合适的NT6系统，可能是WinXP或Win2003。
goto office

:winvista
echo.
echo 当前为Windows Vista/2008，开始尝试激活.....
echo.
set Business=YFKBB-PQJJV-G996G-VWGXY-2V3X8
set BusinessN=HMBQG-8H2RH-C77VX-27R82-VMQBT
set Enterprise=VKK3X-68KWM-X2YGT-QR4M6-4BWMV
set EnterpriseN=VTC42-BM838-43QHV-84HX6-XJXKV
set ServerWeb=WYR28-R7TFJ-3X2YQ-YCY4H-M249D
set ServerStandard=TM24T-X9RMF-VWXK6-X8JC9-BFGM2
set ServerStandardV=W7VD6-7JFBR-RX26B-YKQ3Y-6FFFJ
set ServerEnterprise=YQGMW-MPWTJ-34KDK-48M3W-X4Q6V
set ServerEnterpriseV=39BXF-X8Q23-P2WWT-38T2F-G3FPG
set ServerWeb=RCTX3-KWVHP-BR6TB-RB6DM-6X7HP
set ServerDatacenter=7M67G-PC374-GR742-YH8V4-TCBY3
set ServerDatacenterV=22XQ2-VRXRG-P8D42-K34TD-G3QQC
set ServerEnterpriseIA64=4DWFP-JF3DJ-B7DTH-78FJB-PDRHK
goto windowsstart

:win7
echo.
echo 当前为Windows 7/2008 R2，开始尝试激活.....
echo.
set Professional=FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4
set Professional=GMJQF-JC7VC-76HMH-M4RKY-V4HX6
set ProfessionalN=MRPKT-YTG23-K7D7T-X2JMM-QY7MG
set ProfessionalE=W82YF-2Q76Y-63HXB-FGJG9-GF7QX
set Enterprise=33PXH-7Y6KF-2VJC9-XBBR8-HVTHH
set EnterpriseN=YDRBP-3D83W-TY26F-D46B2-XCKRJ
set EnterpriseE=C29WB-22CC8-VJ326-GHFJW-H9DH4
set Ultimate=49PB6-6BJ6Y-KHGCQ-7DDY6-TF7CD
set HomePremium=CQBVJ-9J697-PWB9R-4K7W4-2BT4J
set HomeBasic=2P6PB-G7YVY-W46VJ-BXJ36-PGGTG
set Starter=273P4-GQ8V6-97YYM-9YTHF-DC2VP
set ServerWeb=6TPJF-RBVHG-WBW2R-86QPH-6RTM4
set ServerHPC=TT8MH-CG224-D3D7Q-498W2-9QCTX
set ServerStandard=YC6KT-GKW9T-YTKYR-T4X34-R7VHC
set ServerEnterprise=489J6-VHDMP-X63PK-3K798-CPX3Y
set ServerDatacenter=74YFP-3QFB3-KQT8W-PMXWJ-7M648
set ServerEnterpriseIA64=GT63C-RJFQ3-4GMB6-BRFB9-CB83V
goto windowsstart
:win8
echo.
echo 当前为Windows 8/2012，开始尝试激活.....
echo.
set Professional=NG4HW-VH26C-733KW-K6F98-J8CK4
set ProfessionalN=XCVCF-2NXM9-723PB-MHCB7-2RYQQ
set Core=BN3D2-R7TKB-3YPBD-8DRP2-27GG4
set Enterprise=32JNW-9KQ84-P47T8-D8GGY-CWCK7
set EnterpriseN=JMNMF-RHW7P-DMY6X-RF3DR-X2BQT
set CoreN=8N2M2-HWPGY-7PGT9-HGDD8-GVGGY
set CoreSingleLanguage=2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
set CoreCountrySpecific=4K36P-JN4VD-GDC6V-KDT89-DYFKP
set ServerMultiPointPremium=XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G
set ServerMultiPointStandard=HM7DN-YVMH3-46JC3-XYTG7-CYQJJ
set ServerStandard=XC9B7-NBPP2-83J2H-RHMBY-92BT4
set ServerDatacenter=48HP8-DN98B-MYWDG-T2DCC-8W83P
set ProfessionalWMC=GNBB8-YVD74-QJHX6-27H4K-8QHDG
set CoreARM=DXHJF-N9KQX-MFPVR-GHGQK-Y7RKV
goto windowsstart
:win81
echo.
echo 当前为Windows 8.1/2012R2，开始尝试激活.....
echo.
set Core=M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK
set CoreARM=XYTND-K6QKT-K2MRH-66RTM-43JKP
set CoreCountrySpecific=NCTT7-2RGK8-WMHRF-RY7YQ-JTXG3
set CoreN=7B9N3-D94CG-YTVHR-QBPX3-RJP64
set CoreSingleLanguage=BB6NG-PQ82V-VRDPW-8XVD2-V8P66
set EmbeddedIndustry=NMMPB-38DD4-R2823-62W8D-VXKJB
set EmbeddedIndustryA=VHXM3-NR6FT-RY6RT-CK882-KW2CJ
set EmbeddedIndustryE=FNFKF-PWTVT-9RC8H-32HB2-JB34X
set Enterprise=MHF9N-XY6XB-WVXMC-BTDCT-MKKG7
set EnterpriseN=TT4HM-HN7YT-62K67-RGRQJ-JFFXW
set Professional=GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
set ProfessionalN=HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
set ProfessionalWMC=789NJ-TQK6T-6XTH8-J39CJ-J8D3P
set ServerCloudStorageCore=3NPTF-33KPT-GGBPR-YX76B-39KDD
set ServerCloudStorage=3NPTF-33KPT-GGBPR-YX76B-39KDD
set ServerDatacenter=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
set ServerDatacenterCore=W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9
set ServerStandard=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
set ServerStandardCore=D2N9P-3P6X9-2R39C-7RTCD-MDVJX
set ServerSolution=KNC87-3J2TX-XB4WP-VCPJV-M4FWM
set ServerSolutionCore=KNC87-3J2TX-XB4WP-VCPJV-M4FWM
goto windowsstart
:win10
echo.
echo. & echo 当前为Windows 10或Windows 11，开始尝试激活.....
echo.
set Core=TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
set CoreCountrySpecific=PVMJN-6DFY6-9CCP6-7BKTT-D3WVR
set CoreN=3KHY7-WNT83-DGQKR-F7HPR-844BM
set CoreSingleLanguage=7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
set ProfessionalStudent=YNXW3-HV3VB-Y83VG-KPBXM-6VH3Q
set ProfessionalStudentN=8G9XJ-GN6PJ-GW787-MVV7G-GMR99
set Professional=W269N-WFGWX-YVC9B-4J6C9-T83GX
set Professional=NXRQM-CXV6P-PBGVJ-293T4-R3KTY
set Professional=N7CXQ-RW8G8-QT9RC-9BBRH-YY49M
set Professional=YC7N8-G7WR6-9WR4H-6Y2W4-KBT6X
set ProfessionalN=MH37W-N47XK-V7XM9-C7227-GCQG9
set ProfessionalSN=8Q36Y-N2F39-HRMHT-4XW33-TCQR4
set ProfessionalWMC=NKPM6-TCVPT-3HRFX-Q4H9B-QJ34Y
set ProfessionalWorkstation=NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J
set ProfessionalWorkstationN=9FNHH-K3HBT-3W4TD-6383H-6XYWF
set ProfessionalEducation=6TP4R-GNPTD-KYYHQ-7B7DP-J447Y
set ProfessionalEducationN=YVWGF-BXNMC-HTQYQ-CPQ99-66QFC
set Education=NW6C2-QMPVW-D7KKK-3GKT6-VCFB2
set Education=YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY
set EducationN=2WH4N-8QGBV-H22JP-CT43Q-MDWWJ
set Enterprise=NPPR9-FWDCX-D2C8J-H872K-2YT43
set Enterprise=VY72Q-T3NYM-MJ2MJ-9M8Q9-M98WR
set Enterprise=DF4VN-TGK62-CC8YB-CDQ2H-HXMPF
set EnterpriseN=DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4
set EnterpriseS=WNMTR-4C88C-JK8YV-HQ7T2-76DF9
set EnterpriseSN=2F77B-TNFGY-69QQF-B8YKP-D69TJ
set EnterpriseN=DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ
set EnterpriseSN=QFFDN-GRT3P-VKWWX-X7T3R-8B639
set EnterpriseG=YYVX9-NTFWV-6MDM3-9PT4T-4M68B
set EnterpriseGN=44RPN-FTY23-9VTTB-MP9BX-T84FV

goto windowsstart
:windowsstart
for /f "tokens=3 delims= " %%i in ('reg QUERY "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v "EditionID"') do set EditionID=%%i
if defined %EditionID% (
        cscript //Nologo %windir%\system32\slmgr.vbs /ipk !%EditionID%!
        cscript //Nologo %windir%\system32\slmgr.vbs /skms kms.zbgw.xyz
        cscript //Nologo %windir%\system32\slmgr.vbs /ato
) else (
        echo 找不到系列号，可能是旗舰版之类的系统……
)
goto office
:wintooff
echo 系统已经永久激活！
echo 下面开始转入Office激活。
goto office
:office
echo 检查安装的office……
call :GetOfficePath 14 Office2010
call :ActOffice 14 Office2010
call :GetOfficePath 15 Office2013
call :ActOffice 15 Office2013
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles%\Microsoft Office\Office16
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" set _Office16Path=%ProgramFiles(x86)%\Microsoft Office\Office16
if DEFINED _Office16Path (echo.&echo 已发现 Office2016
    call :ActOffice 16 Office2016
  ) else (echo.&echo 未发现 Office2016)
echo 操作成功完成！
start "" slmgr.vbs -dlv
echo.&pause
exit

:ActOffice
if DEFINED _Office%1Path (
    cd /d "!_Office%1Path!"
        echo 检查 %2 的激活状态。
        cscript //nologo ospp.vbs /act | find /i "successful" > NUL && (
        echo.&echo  %2 已经激活，自动跳过  & echo.) || (
                echo.&echo  %2 未激活，正尝试进行激活。
                if %1 EQU 16 call :Licens16
                cscript //nologo ospp.vbs /sethst:kms.zbgw.xyz  >nul
                cscript //nologo ospp.vbs /act | find /i "successful" && (
        echo.&echo ***** %2 激活成功 ***** & echo.) || (echo.&echo ***** %2 激活失败 ***** & echo.)
                )
)   
cd /d "%~dp0"
goto :EOF

:GetOfficePath
echo.&echo 正在检测 %2 系列产品的安装路径.....
set _Office%1Path=
set _Reg32=HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\%1.0\Common\InstallRoot
set _Reg64=HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\%1.0\Common\InstallRoot
reg query "%_Reg32%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg32%" /v "Path"') do SET "_OfficePath1=%%b"
reg query "%_Reg64%" /v "Path" > nul 2>&1 && FOR /F "tokens=2*" %%a IN ('reg query "%_Reg64%" /v "Path"') do SET "_OfficePath2=%%b"
if DEFINED _OfficePath1 (if exist "%_OfficePath1%ospp.vbs" set _Office%1Path=!_OfficePath1!)
if DEFINED _OfficePath2 (if exist "%_OfficePath2%ospp.vbs" set _Office%1Path=!_OfficePath2!)
set _OfficePath1=
set _OfficePath2=
if DEFINED _Office%1Path (echo.&echo 已发现 %2) else (echo.&echo 未发现 %2)
goto :EOF


:Licens16
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\standardvl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_kms*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\project???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\proplusvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\standardvl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
for /f %%x in ('dir /b ..\root\Licenses16\visio???vl_mak*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%x" >nul
cscript ospp.vbs /inpkey:NYH39-6GMXT-T39D4-WVXY2-D69YY >nul
cscript ospp.vbs /inpkey:7WHWN-4T7MP-G96JF-G33KR-W8GF4 >nul
cscript ospp.vbs /inpkey:RBWW7-NTJD4-FFK2C-TDJ7V-4C2QP >nul
cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >nul
cscript ospp.vbs /inpkey:JNRGM-WHDWX-FJJG3-K47QV-DRTFM >nul
cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT >nul
cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK >nul
cscript ospp.vbs /inpkey:HFTND-W9MK4-8B7MJ-B6C4G-XQBR2 >nul
cscript ospp.vbs /inpkey:GNFHQ-F6YQM-KQDGJ-327XX-KQBVC >nul
cscript ospp.vbs /inpkey:GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW >nul
cscript ospp.vbs /inpkey:9C2PK-NWTVB-JMPW8-BFT28-7FTBF >nul
cscript ospp.vbs /inpkey:DR92N-9HTF2-97XKM-XW2WJ-XW3J6 >nul
cscript ospp.vbs /inpkey:R69KK-NTPKF-7M3Q4-QYBHW-6MT9B >nul
cscript ospp.vbs /inpkey:RJ7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6 >nul
cscript ospp.vbs /inpkey:F47MM-N3XJP-TQXJ9-BP99D-8K837 >nul
cscript ospp.vbs /inpkey:869NQ-FJ69K-466HW-QYCP2-DDBV6 >nul
cscript ospp.vbs /inpkey:WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6 >nul
cscript ospp.vbs /inpkey:WGT24-HCNMF-FQ7XH-6M8K7-DRTW9 >nul
cscript ospp.vbs /inpkey:D8NRQ-JTYM3-7J2DX-646CT-6836M >nul
cscript ospp.vbs /inpkey:69WXN-MBYV6-22PQG-3WGHK-RM6XC >nul
cscript ospp.vbs /inpkey:NY48V-PPYYH-3F4PX-XJRKJ-W4423 >nul
cscript ospp.vbs /inpkey:DMTCJ-KNRKX-26982-JYCKT-P7KB6 >nul
goto :EOF

exit
:xptooff
echo 当前为WinXP或Win2003。
call :GetOfficePath 14 Office2010
call :ActOffice 14 Office2010


:fail
cls
echo.
echo 无法连接到服务器，请编辑脚本更换KMS服务器，或使用其它激活工具.....
pause
exit


