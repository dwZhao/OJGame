# imports
import xlrd
import xlsxwriter
import time
import copy

NUM_GROVES = 6
GROVE_NAMES = ["FLA", "CAL", "TEX", "ARZ", "BRA", "SPA"]
PRODUCT_NAMES = ["ORA", "POJ", "ROJ", "FCOJ"]
FUTURES_NAMES = ["ORA", "FCOJ"]
LB_PER_TON = 2000
SHIP_COST_GPS = 0.22
SHIP_COST_SM = 1.2

def getOpenPlants(sheet):
	"Get list of open processing plants"

	plants = []

	for i in range(0, 10):
		if sheet.cell_value(5 + i, 3) != 0:
			plants.append(sheet.cell_value(5 + i, 1))

	return plants

def getOpenStorages(sheet):
	"Get list of open storages"

	storages = []

	for i in range(0, 71):
		if sheet.cell_value(35 + i, 3) != 0:
			storages.append(sheet.cell_value(35 + i, 1))		

	return storages

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
	"Get absolute amount of futures arriving each month each week"
	Mat = getMatFutures(sheet)
	percentages = getFuturesArrivalPercentage(sheet)

	return {future:[(per/100) * Mat[future] for per in percentages[future]] for future in FUTURES_NAMES}

def getGPSDistance(sheet, openPlants, openStorages):
	"Get distance from groves to plants and storages"

	GPSDist = {}

	for counter, grove in enumerate(GROVE_NAMES):
		distances = {}
		for i in range(0, 81):
			facility = sheet.cell_value(1 + i, 0)
			if (facility in openStorages) or (facility in openPlants):
				distances[facility] = sheet.cell_value(1 + i, 1 + counter)
		GPSDist[grove] = distances

	return GPSDist

def getFCOJTransportCosts(sheet, amountMatFutures, GPSdist, openStorages):
	"Calculate FCOJ transportation costs"

	transportCosts = {}
	for counter, storage in enumerate(openStorages):
		costs = []
		for month in range(0, 12):
			costs.append(SHIP_COST_GPS * (sheet.cell_value(26 + counter, 2) / 100) * amountMatFutures["FCOJ"][month] * GPSdist["FLA"][sheet.cell_value(26 + counter, 1)])
		transportCosts[storage] = costs

	return transportCosts

def getFCOJTransportQuantity(sheet, amountMatFutures, openStorages):
	"Calculate FCOJ shipped to facilities"

	transportQuantities = {}
	for counter, storage in enumerate(openStorages):
		quantities = []
		for month in range(0, 12):
			quantities.append((sheet.cell_value(26 + counter, 2) / 100) * amountMatFutures["FCOJ"][month])
		transportQuantities[storage] = quantities

	return transportQuantities

def getActualGroveAmount(orderQuantities, amountMatFutures):
	"Get total amount of oranges shipped out of groves, including futures"

	originalOrders = copy.deepcopy(orderQuantities)
	newOrder = []

	for month in range(0, 12):
		newOrder.append(tuple(order + amountMatFutures["ORA"][month] for order in originalOrders["FLA"][month]))

	originalOrders["FLA"] = newOrder

	return originalOrders

def getGroveORAShipPer(sheet, openPlants, openStorages):
	"Get shipping breakdown from grove to plants and storages"

	shipPer = {}

	for counter, grove in enumerate(GROVE_NAMES):
		percentages = {}
		for i, facility in enumerate(openPlants + openStorages):
			percentages[facility] = sheet.cell_value(5 + counter, 2 + i) / 100
		shipPer[grove] = percentages

	#print(shipPer)
	return shipPer

def getGroveORAShipCost(sheet, actualGroveAmount, GPSDist, openPlants, openStorages):
	"Get the cost of shipping ORA and ORA futures to plants and storages"

	groveORAShipPer = getGroveORAShipPer(sheet, openPlants, openStorages)

	shipCost = {}

	for grove in GROVE_NAMES:
		#print(grove)
		facilityShipCost = {}
		for facility in (openPlants + openStorages):
			#print(facility)
			dist = GPSDist[grove][facility]
			per = groveORAShipPer[grove][facility]
			cost = []
			for month in range(0, 12):
				cost.append(tuple(SHIP_COST_GPS * ORA * dist * per for ORA in actualGroveAmount[grove][month]))
			facilityShipCost[facility] = cost
		shipCost[grove] = facilityShipCost

	return shipCost

def getGroveORAShipQuantities(sheet, actualGroveAmount, openPlants, openStorages):
	"Get amounts shipped from groves to facilities"

	groveORAShipPer = getGroveORAShipPer(sheet, openPlants, openStorages)

	shipQuantity = {}

	for grove in GROVE_NAMES:
		facilityShipQuantity = {}
		for facility in (openPlants + openStorages):
			per = groveORAShipPer[grove][facility]
			quantity = []
			for month in range(0, 12):
				quantity.append(tuple(ORA * per for ORA in actualGroveAmount[grove][month]))
			facilityShipQuantity[facility] = quantity
		shipQuantity[grove] = facilityShipQuantity

	return shipQuantity

def getStorageMarketPref(sheet, openStorages):
	"Get the nearest storage to each market"

	pref = {}

	for index, market in enumerate(sheet.col_values(1, 1, 101)):
		minDist = float("inf")
		cellIndex = float("inf")

		for i in range(0, 71):
			if sheet.cell_value(0, i + 2) in openStorages:
				dist = sheet.cell_value(index + 1, i + 2)
				if minDist > dist:
					minDist = dist
					cellIndex = i + 2
		pref[market] = [sheet.cell_value(0, cellIndex), minDist, SHIP_COST_SM * minDist]
	return pref

def main():
	#decBook = xlrd.open_workbook("decisionSheet.xlsx")
	#decBook = xlrd.open_workbook('thebreakfastclub2014.xlsx')
	decBook = xlrd.open_workbook('2014.xlsx')

	exoBook = xlrd.open_workbook("Exo.xlsx")
	distBook = xlrd.open_workbook("StaticDataMod.xlsx")

	rawSheet = decBook.sheet_by_name("raw_materials")
	facSheet = decBook.sheet_by_name("facilities")
	shipManuSheet = decBook.sheet_by_name("shipping_manufacturing")
	groveSheet = exoBook.sheet_by_name("Grove")
	GPSSheet = distBook.sheet_by_name("G->PS")
	SMSheet = distBook.sheet_by_name("S->M")

	harvestPrices = getHarvestPrices(groveSheet)
	harvestQuantities = getHarvestQuantities(groveSheet)
	orderQuantities = getOrderQuantities(rawSheet, harvestPrices, harvestQuantities)

	orderCost = getOrderCost(harvestPrices, orderQuantities)

	amountMatFutures = getFuturesArrivalAmount(rawSheet)

	openPlants = getOpenPlants(facSheet)
	openStorages = getOpenStorages(facSheet)

	GPSDist = getGPSDistance(GPSSheet, openPlants, openStorages)

	FCOJTransportCost = getFCOJTransportCosts(shipManuSheet, amountMatFutures, GPSDist, openStorages)

	actualGroveAmount = getActualGroveAmount(orderQuantities, amountMatFutures)

	#print(GPSDist)

	shipCost = getGroveORAShipCost(shipManuSheet, actualGroveAmount, GPSDist, openPlants, openStorages)

	#getGroveORAShipPer(shipManuSheet, openPlants, openStorages)


	for grove in GROVE_NAMES:
		print(grove + " " + str(getGroveORAShipQuantities(shipManuSheet, actualGroveAmount, openPlants, openStorages)[grove]))

	pref = getStorageMarketPref(SMSheet, openStorages)

	# print(getGroveORAShipQuantities(shipManuSheet, actualGroveAmount, openPlants, openStorages))

if __name__ == "__main__":
    main()