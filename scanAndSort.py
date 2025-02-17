import pyttsx3
import info
engine = pyttsx3.init()

allISBNs = []
isbn = input("Scan a book and press enter: ")
while isbn:
    if isbn.lower().startswith("c"): # To Cancel the previous scan
        if allISBNs: allISBNs.pop()
        continue

    try:
        retailData = info.getISBNRetailData(isbn)
    except ValueError as e:
        engine.say("Please rescan")
        engine.runAndWait()
        isbn = input("Scan a book and press enter: ")
        continue
    allISBNs.append(isbn)
    topSeller = retailData[0]["retailer"]
    print(topSeller)
    engine.say(f"{topSeller} for {retailData[0]['price']}")
    engine.runAndWait()
    isbn = input("Scan a book and press enter: ")


for book in allISBNs[:-1]:
    print(book, end=", ")

print(allISBNs[-1])

print(f"Total books scanned: {len(allISBNs)}")






