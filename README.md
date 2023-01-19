# CMScrape
CardMarket Scraper to fetch MTG card prices and keep track of price changes inside a .xlsx file.

---
## Requirements & Notes
Python 3.10

Windows 10 (no other OS were tested)

Microsoft Excel (no other similar tools were tested)

This version was only tested on MTG cards on cardmarket.com with url's leading to a page which is structured like the one in the link below.
This version is working as of 19.01.2023. This may or may not be the case in the future and is subject to changes made to this repository or
to the page structure of cardmarket.com.


## Setup & Usage Guide
Download the zip file and unzip it in any folder of choice. 

Open `./data/urls.txt` with any editor, copy in url's like the following one and hit save. Keep one url in each line.
[https://www.cardmarket.com/en/Magic/Cards/Demonic-Tutor?sellerCountry=7&isPlayset=N](https://www.cardmarket.com/en/Magic/Cards/Demonic-Tutor?sellerCountry=7&isPlayset=N)

In case there is no `./data/cardprices.xlsx` inside the folder simply create one inside there with the exact same name.

Open `cmd` or `PowerShell` and change the working directory with `cd` to the folder the zip was extracted to.
After that run the following two code lines. The first one installs all necessary packages. The second one executes the python script.
```
$ python -m pip install -r requirements.txt
$ python CMScrape.py
```
For each url in `./data/urls.txt` a completion/failed message will be displayed in the `cmd` or `PowerShell` window. Once every url has been
parsed a task completed will be shown.


## Example urls.txt input and cardprices.xlsx output
![alt text](https://i.imgur.com/LVU6rjQ.png)
![alt text](https://i.imgur.com/0vv8ahd.png)

