import requests
retailerInformation = {
    "Worthless Doorstops/Paperweights": {
        "minimumOrder": 1,
        "id": 0,
        "allowsBulk": True
    }
}

def getISBNRetailData(isbn):
    """
    Returns a list of dictionaries containing the retail data for a given ISBN
    sorted by selling price.
    """
    global retailerInformation
    data = requests.get(f"https://api.bookscouter.com/v4/prices/sell/{isbn}")
    json = data.json()
    retailData = []
    for price in json["prices"]:
        if price["price"] == 0:
            break
        retailData.append({
            "isbn": isbn,
            "title": json["book"]["title"],
            "price": price["price"],
            "retailer": price["vendor"]["name"],
            "minimumOrder": price["vendor"]["minimumOrder"],
            "imageURL": json["book"]["image"].replace("SL75", "SL3000"),
            "retailerURL": f"https://api.bookscouter.com/exits/sell/{price['vendor']['id']}/{isbn}",
        })
        retailerInformation[price["vendor"]["name"]] = {"minimumOrder": price["vendor"]["minimumOrder"],
                                                        "id": price["vendor"]["id"],
                                                        "allowsBulk": price["vendor"]["bulkInfo"]["allowBulk"]
                                                        }

    retailData = retailData or [{
        "isbn": isbn, 
        "title": json["book"]["title"], 
        "price": 0, 
        "retailer": "Worthless Doorstops/Paperweights", 
        "minimumOrder": 1,
        "imageURL": "",
        "retailerURL": "",
        }]
    return sorted(retailData, reverse=True, key=lambda e: e["price"])



