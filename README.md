# Python Insert Photo To Doc

*Reference Link*:[工料測量日常(QS)](https://maincontractorqs.wordpress.com/2022/07/26/script-share-%e6%94%be%e7%9b%b8%e5%85%a5word-%e6%aa%94/)

According to the reference link, I have modified the table cell. 

In addition, it can support photo file name with different characters such as chinese.

## ***Step:***
1. Run "***insert_file.exe***" application.
2. Select Folder which includes a photo that you would like to insert into a document.
3. Find your doc file according to your selected folder path.


In addition, you can modify the "***insert_file.py***" for further development.

Please follow the link below about how to create an executable of Python Script for your reference:
https://datatofish.com/executable-pyinstaller/

If you don't even install the pyinstaller, you can run it on command
```
pip install pyinstaller

```
Then, you go to your file's path
```
cd <your python file path>
```
Create the Executable with pyinstaller
```
pyinstaller --onefile <your filename>.py
```
Finally, you'll see the executable file in **dist** folder.

:+1: Good Luck!~
