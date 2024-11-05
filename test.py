import time
import requests

while True:
    # Fetch the current price of Bitcoin in USD
    url = "https://api.coindesk.com/v1/bpi/currentprice/USD.json"
    response = requests.get(url, headers={"X-CoinDesk-APIKey": API_KEY})
    data = response.json()

    # Print the current price of Bitcoin
    print("The current price of Bitcoin is $" + data["bpi"]["USD"]["rate"])

    # Pause for 5 minutes
    time.sleep(300)