# pl-bonds-analyzer

## How to use
1. Download the latest version of this script
2. Download monthly statistics and copy to ./data/dane.xls
3. Run the script
```python3 script.py```

## Paramters:
- only active bonds turnover threshold
- risk threshold relative to average margin
- time to maturity
- best time to buy (Todo)
- issuers exclude list ( SP, Miasto, Gmina ) (Todo)
- only active bonds transaction threshold (Todo)

## Reverse engineering
Get month state
https://gpwcatalyst.pl/pub/CATALYST/statystyki/statystyki_miesieczne/202207_CAT.xls

Get margin
curl -X POST https://obligacje.pl/ajax/kalkulatorDane.php -H "Content-Type: application/x-www-form-urlencoded" -d "dane=kal_kod_obligacji%3DCSA0726"