***This is one of the tools I developed for thermal team when I worked during the last year.
***This tool is used for extracting specific data(such as the server's temperature and PWM speed) from a bunch of logs, and generate corresponding Microsoft Excel reports.

The directories of the project are as bellow.
1. doc
	contains the readme.pdf/readme.docx, shows how to use the tool
2. ref
	the original log files and output Microsoft Excel templates
3. release
	the release tool for thermal team
4. thermal
	the kernel code for this tool
5. build.py
	automatically encapsulate the python code to an executable tool along with the readme.pdf to a compressed file in the release directory