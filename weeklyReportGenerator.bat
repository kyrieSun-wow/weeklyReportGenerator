@echo off
set inPutFile=
set outPutFle=WeeklyReport.xlsx
set sheetName=


echo ��ѴӶ���������xlsx�ļ��ϵ��������ڲ��س�:
set /p inPutFile=
echo.

echo ������Խ������ݵ�xlsx�ļ��϶��������ڲ��س�������Ҫ�½�����ʹ�õ�ǰĿ¼�µ�Ĭ���ļ���ֱ�ӻس�����:
set /p outPutFle=
echo.

echo �������µ�sheet���Ʋ��س�����ʾ��MME/AMF/...  ��Ӧ��TeamMateConfig.txt�������õ�team name,���ִ�Сд��:
set /p sheetName=
echo.

Rem echo.&echo.
echo ���ڴ�%outPutFle%���½�%sheetName%���Խ��մ�%inPutFile%�����������...
Rem start python3 %~dp0weeklyReportGenerator.py -i inPutFile -o outPutFle -n sheetName
Rem set cmd_str=title Weekly Report generator v1.0 & python weeklyReportGenerator.py -i %inPutFile% -o %outPutFle% -n %sheetName%
set cmd_str=python weeklyReportGenerator.py -i %inPutFile% -o %~dp0%outPutFle% -n %sheetName%
Rem echo cmd_str is %cmd_str%
Rem cmd /k "title Weekly Report generator v1.0 & python weeklyReportGenerator.py -i %inPutFile% -o %outPutFle% -n %sheetName%"
cmd /k %cmd_str%
Rem echo %outPutFle%
Rem echo %sheetName%
echo.

echo ������ϡ� 

pause