@echo off  
set startime=%date:~,4%-%date:~5,2%-%date:~8,2% %time:~0,8%
echo auto_job.bat ��ʼʱ��:%startime% >> log.txt
python D:\project\project-carNew\priceReport\uploadToOracle.py
python D:\project\project-carNew\priceReport\createReportFile.py
python D:\project\project-carNew\priceReport\markValue.py
python D:\work\priceReport\sendEmail.py
set endtime=%date:~,4%-%date:~5,2%-%date:~8,2% %time:~0,8%
echo auto_job.bat ����ʱ��:%endtime% >> log.txt
pause