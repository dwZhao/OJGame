# imports
import xlrd
import xlsxwriter
import time

NUM_GROVES = 6
GROVE_NAMES = ["FLA", "CAL", "TEX", "ARZ", "BRA", "SPA"]
PRODUCT_NAMES = ["ORA", "POJ", "ROJ", "FCOJ"]
FUTURES_NAMES = ["ORA", "FCOJ"]
LB_PER_TON = 2000

def getExchangeRates(sheet):
	"Get exchange rates for Brazil and Spain"

	exchangeRates = {}

	for counter, grove in enumerate(GROVE_NAMES[4:]):
		rates = sheet.row_values(13 + counter, 2, 14)
		exchangeRates[grove] = rates

	return exchangeRates

def getHarvestPrices(sheet):
	"Get harvest prices for the year by month by region"

	harvestPrices = {}
	exchangeRates = getExchangeRates(sheet)
	# Put harvest prices from exo info into list and then dict with grove as key
	for counter, grove in enumerate(GROVE_NAMES):
		prices = sheet.row_values(4 + counter, 2, 14)

		if grove in GROVE_NAMES[4:]:
			harvestPrices[grove] = [r * p for r, p in zip(exchangeRates[grove], prices)]
		else:
			harvestPrices[grove] = prices

	return harvestPrices

def getMultipliers(sheet):
	"Get order multipliers"

	multipliers = {}

	for counter, grove in enumerate(GROVE_NAMES):
		groveMultipliers = []
		for i in range(0, 3):
			groveMultipliers.append((sheet.cell_value(16 + counter, 2 + 2 * i), sheet.cell_value(16 + counter, 3 + 2 * i)))
		multipliers[grove] = groveMultipliers

	return multipliers

def getHarvestQuantities(sheet):
	"Get harvest quantities that limit orders"

	harvestQuantities = {}

	for counter, grove in enumerate(GROVE_NAMES):
		groveQuantities = []
		for i in range(0, 12):
			groveQuantities.append(tuple(sheet.row_values(37 + counter, 2 + 4 * i, 6 + 4 * i)))

		harvestQuantities[grove] = groveQuantities

	return harvestQuantities

def getOrderQuantities(sheet, harvestPrices, harvestQuantities):
	"Get order quantities for the year by month by region"

	multipliers = getMultipliers(sheet)

	orderQuantities = {}
	
	for counter, grove in enumerate(GROVE_NAMES):
		requestedOrders = sheet.row_values(5 + counter, 2, 14)
		monthlyOrders = []
		for month in range(0, 12):
			if (harvestPrices[grove][month] < multipliers[grove][0][1]):
				requestedOrders[month] *= multipliers[grove][0][0]
			elif (harvestPrices[grove][month] < multipliers[grove][1][1]):
				requestedOrders[month] *= multipliers[grove][1][0]
			elif (harvestPrices[grove][month] < multipliers[grove][2][1]):
				requestedOrders[month] *= multipliers[grove][2][0]
			else:
				requestedOrders[month] = 0
			monthTuple = []
			for week in range(0, 4):
				if (requestedOrders[month] > harvestQuantities[grove][month][week]):
					monthTuple.append(harvestQuantities[grove][month][week])
				else:
					monthTuple.append(requestedOrders[month])
			monthlyOrders.append(tuple(monthTuple))
		orderQuantities[grove] = monthlyOrders

	return orderQuantities

def getOrderCost(harvestPrices, orderQuantities):
	"Multiply prices and orders for costs for current year orders"

	orderCost = {}

	for grove in GROVE_NAMES:
		cost = []
		for month in range(0, 12):
			cost.append(tuple(LB_PER_TON * harvestPrices[grove][month] * w for w in orderQuantities[grove][month]))
		orderCost[grove] = cost
	return orderCost

def getMatFutures(sheet):
	"Get matured futures arriving per week in FLA"

	return {FUTURES_NAMES[0]:sheet.cell_value(29, 15), FUTURES_NAMES[1]:sheet.cell_value(35, 15)}

def getFuturesArrivalPercentage(sheet):
	"Get percentage of ORA and FCOJ futures shipped each month to FLA"

	return {FUTURES_NAMES[0]:sheet.row_values(46, 2, 14), 
	        FUTURES_NAMES[1]:sheet.row_values(47, 2, 14)}

def getFuturesArrivalAmount(sheet):
	Mat = getMatFutures(sheet)
	percentages = getFuturesArrivalPercentage(sheet)

	return {future:[(per/100) * Mat[future] for per in percentages[future]] for future in FUTURES_NAMES}
def main():
	decBook = xlrd.open_workbook("decisionSheet.xlsx")
	exoBook = xlrd.open_workbook("Exo.xlsx")

	rawSheet = decBook.sheet_by_name("raw_materials")
	groveSheet = exoBook.sheet_by_name("Grove")

	harvestPrices = getHarvestPrices(groveSheet)
	harvestQuantities = getHarvestQuantities(groveSheet)
	orderQuantities = getOrderQuantities(rawSheet, harvestPrices, harvestQuantities)

	orderCost = getOrderCost(harvestPrices, orderQuantities)

	print(getFuturesArrivalAmount(rawSheet))

	for grove in GROVE_NAMES:
		print(grove + " " + str(orderCost[grove]))

if __name__ == "__main__":
    main()