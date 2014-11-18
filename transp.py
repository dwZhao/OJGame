from pulp import *
import xlrd

sources = ["FLA", "CAL", "TEX", "ARZ"]

supply = {"FLA": 1500,
          "CAL": 1500,
          "TEX": 1500,
          "ARZ": 1500}

sinks = ["P01", "P05", "P07", "Dummy"]

demand = {"P01": 345,
		  "P05": 500,
		  "P07": 3589,
		  "Dummy": 1566}

costs = [ # P01 P05 P07 Dummy
           [454, 2589, 1161, 0], #FLA
           [2669, 351, 2190, 0], #CAL
           [1277, 1211, 1154, 0], #TEX
           [2321, 88, 1181, 0]  #ARZ
        ]

costs = makeDict([sources, sinks], costs, 0)

prob = LpProblem("Network", LpMinimize)

links = [(orig, dest) for orig in sources for dest in sinks]

variables = LpVariable.dicts("Link", (sources, sinks), 0, None, LpInteger)

prob += lpSum([variables[orig][dest] * costs[orig][dest] for (orig, dest) in links]), "Sum of Costs"

for orig in sources:
	prob += lpSum([variables[orig][dest] for dest in sinks]) <= supply[orig], "Sum of source flow %s" %orig

for dest in sinks:
	prob += lpSum([variables[orig][dest] for orig in sources]) >= demand[dest], "Sum of sink flow %s"%dest

prob.writeLP("transportation.lp")

prob.solve()

print "Status:", LpStatus[prob.status]

for v in prob.variables():
    print v.name, "=", v.varValue
   
print "Total Cost of Transportation = ", value(prob.objective)