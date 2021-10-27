
# Bean Seller

A New Worlds market scraper.


## Deployment
This must be installed on Windows as it uses the Windows api to do its stuff

### Install Prerequisites: 

Pytesseract:   
Good writeup here: https://medium.com/@marioruizgonzalez.mx/how-install-tesseract-orc-and-pytesseract-on-windows-68f011ad8b9b
   
pywin32:
https://github.com/mhammond/pywin32


  
## Run Locally

Clone the project

```bash
  git clone https://link-to-project
```

Go to the project directory

```bash
  cd my-project
```

Install dependencies

```bash
  pip install -r requirements.txt
```

Start the scraper

```bash
  python .\main.py
```

When closed, by pressing q when prompted, this will crate a .json file with its scrapped contents.

  
## Usage

    1. Start New World, go to a market.
    2. Open the market and select what you want scraped (select all markets, all items.)
    3. Start the scrapper.
    4. Press any key when prompted.
    5. Scroll the market list to get the next section of items.
    6. Press any key again.
    7. Scroll all the way down to the bottom of the list.
    8. Press 'w' to get the last two items. 
    9. Go to the next page of the market.
    10. Start this list over from 4.

  
## Things to do (that I probably wont).

Automate the manipulation of the market screen to make this completely hands off.

## What to do with the json output?

I like to put it in https://app.rawgraphs.io/ and figure out what's going on in the markets.