>### Simple utility To keep log of your day-today work.

- Pre-requisite is to type command: "pip install openpyxl"
- Python GUI based tool which logs your work in simple excelsheet file.
- Helps in tracking work history.[or watever you log] 

Known Issues:

- Today, I tried running it on Ubuntu 16 and realised .pyw extension was not working with linux platform.
- .ico has some issue in loading and I dont have time to fix now. So, I am doing some quick changes so that a 'Ubuntu' user can quickly use this tool.
- If you are on 'Windows', you can checkout previous commit.


To create single executable:

- install 'pyinstaller' using pip

	> pyinstaller workipian.py --onefile

 This created a 'Single Executable' file which you can directly use by double clicking.

