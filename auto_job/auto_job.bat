@echo off  
set startime=%date:~,4%-%date:~5,2%-%date:~8,2% %time:~0,8%
echo auto_job.bat 开始时间:%startime% >> log.txt
python D:\project\project-carNew\priceReport\uploadToOracle.py
python D:\project\project-carNew\priceReport\createReportFile.py
python D:\project\project-carNew\priceReport\markValue.py
python D:\work\priceReport\sendEmail.py
set endtime=%date:~,4%-%date:~5,2%-%date:~8,2% %time:~0,8%
echo auto_job.bat 结束时间:%endtime% >> log.txt
pause