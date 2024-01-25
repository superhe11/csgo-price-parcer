# csgo-price-parcer
Basic python script for parcing prices from Steam market.
## Quick How-To-Guide
1. Install Chrome (if you haven't yet)
2. Download ChromeDriver from official page for the google version you have
3. Put chromedriver.exe in THE SAME folder, as python script.
4. Install python
5. Open cmd, go to the directory with script (and .txt file) and run pip install -r requirements.txt
6. Open code by writing in cmd python csparcer.py
   
## Important notes
1. There is a limit on requests to the Steam market, the code automatically will sent you on a cooldown in most cases, but if you see too many timeout errors - just wait ~5 min and restart the script.
2. Script automatically creates and updates excel file please DO NOT rename it.

## Troubleshooting
1. If there is some errors like: ERROR: Couldn't read tbsCertificate as SEQUENCE, ERROR: Failed parsing Certificate etc. chill, they wont affect the code.
2. Sometimes items from market wont parse, they will throw "exceptions.TimeoutException" -> code will just skip this item, i can't do much about it, because its issues on serverside
3. If the code doesn't start at all - try installing all libralies manually: pip install webdriver_manager==3.4.2 openpyxl==3.0.17 psutil==5.8.0 selenium==3.141.0 beautifulsoup4==4.10.0
  
