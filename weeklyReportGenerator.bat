@echo off
set inPutFile=
set outPutFle=WeeklyReport.xlsx
set sheetName=


echo 请把从钉钉导出的xlsx文件拖到到窗口内并回车:
set /p inPutFile=
echo.

echo 请把用以接收数据的xlsx文件拖动到窗口内并回车，若需要新建或者使用当前目录下的默认文件则直接回车即可:
set /p outPutFle=
echo.

echo 请输入新的sheet名称并回车，提示（MME/AMF/...  对应于TeamMateConfig.txt里面配置的team name,区分大小写）:
set /p sheetName=
echo.

Rem echo.&echo.
echo 正在从%outPutFle%中新建%sheetName%表以接收从%inPutFile%整理出的数据...
Rem start python3 %~dp0weeklyReportGenerator.py -i inPutFile -o outPutFle -n sheetName
Rem set cmd_str=title Weekly Report generator v1.0 & python weeklyReportGenerator.py -i %inPutFile% -o %outPutFle% -n %sheetName%
set cmd_str=python weeklyReportGenerator.py -i %inPutFile% -o %~dp0%outPutFle% -n %sheetName%
Rem echo cmd_str is %cmd_str%
Rem cmd /k "title Weekly Report generator v1.0 & python weeklyReportGenerator.py -i %inPutFile% -o %outPutFle% -n %sheetName%"
cmd /k %cmd_str%
Rem echo %outPutFle%
Rem echo %sheetName%
echo.

echo 处理完毕。 

pause