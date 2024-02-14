# TimeStation

is a tool to manage employee clock-in & -out times.

It's able to support **up to 10 employees** for free.

Its API is free to use but is limited to 5000 requests per day.
[API documnetation](https://www.mytimestation.com/API.asp)

TimeStation API keys are found in Settings>API Keys

**Please do not commit the API keys.**

Employees sign in through generated QR codes or spefic PINs.

This scrpit fetches employee shift records and moves the clock-in & -out times to specified .xlsx files and sheets.

### Requirements:

* python3
* pip

```
pip install requests
pip install openpyxl
```

### .xlsx stuff

File names should contain employee name

Sheet names need to be in **English acronyms** e.i Mar not Mrz.

### cron command

```
0 7 * * * /usr/bin/python3 /path/to/your/script.py
```

### troubleshooting

Close all open instances of the .xlsx files

Check APIKEY is correct

Check folder paths are correct

### cron command

```
0 7 * * * /usr/bin/python3 /path/to/your/script.py
```

### troubleshooting

Close all open instances of the .xlsx files

Check APIKEY is correct

Check folder paths are correct