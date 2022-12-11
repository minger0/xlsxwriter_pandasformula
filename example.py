''' examples to show some xlsxwriter_pandasformula magic '''
import pandas as pd
import xlsxwriter_pandasformula as x

x.Constant('pi',3.14)
x.Constant('e',2.72)

price = x.View({
	"sheet"  : "price example"
, "anchor" : [0, 0]
, "name"   : "PRICE((size), (vendor, food)) [exampleinput.csv]"
, "value"  : pd.read_csv("exampleinput.csv", index_col=[0], header=[0,1])
}) # (size), (vendor, food)

print(price.value.columns)

total = x.View({
	"sheet"  : "price example"
, "anchor" : [7, 0]
, "name"   : "TOTAL(vendor,food)"
, "value"  : x.Formula({
		"total"  : ["total"]
	, "rows"   : ["total"]
	, "cols"   : price.cols()
	, "vals"   : lambda itotal, ivendor, ifood: f'=SUM({price.ref((x.ALL),(ivendor,ifood))})'
	})
})

minprice = x.View({
	"sheet"  : "price example"
, "anchor" : [1, 7]
, "name"   : "MINPRICE(size,food)"
, "value"  : x.Formula({
		"size"   : x.dom['size']
	, "food"   : x.dom['food']
	, "rows"   : ["size"]
	, "cols"   : ["food"]
	, "vals"   : lambda isize, ifood: f'=MIN({price.ref((isize),("Pizz",ifood))},{price.ref((isize),("Izza",ifood))})'
	})
})

avgprice = x.View({
	"sheet"  : "price example"
, "anchor" : [1, 11]
, "name"   : "AVGPRICE(avgprice,vendor))"
, "value"  : x.Formula({
		"avgprice" : ['average']
	, "vendor" : x.dom['vendor']
	, "rows"   : ["avgprice"]
	, "cols"   : ["vendor"]
	, "vals"   : lambda iavgprice, ivendor: f'=MIN({price.ref((x.ALL),(ivendor,x.ALL))})'
	})
})

#self-referencing example: first define the view without values so that it can be referenced, then simply use set(<value function>) to set values
sizes = x.View({
	"sheet"  : "price example"
, "anchor" : [12,0]
, "name"   : "SIZES(diameter,size))"
, "value"  : x.Formula({
		"diameter" : ['cm','dm']
	, "size"   : x.dom['size']
	, "rows"   : ["diameter"]
	, "cols"   : ["size"]
	})
})

def sizes_formula(idiameter, isize):
	if isize=="S" and idiameter=="cm":
		retval="20"
	elif idiameter=="cm":
		retval="(pi/e)*"+sizes.ref((idiameter),(x.dom["size"][x.dom["size"].index(isize)-1]))
	else:
		retval=sizes.ref(("cm"),(isize))+"/10"
	return "="+retval

sizes.set(sizes_formula)

pricenewPizz = x.View({
	"sheet"  : "new price"
, "anchor" : [0, 0]
, "name"   : "PRICENEWPIZZ((size),(newprice,food)))"
, "value"  : x.Formula({
		"size"   : x.dom['size']
	, "newprice": ["new price Pizz"]
	, "food"   : x.dom["food"]
	, "rows"   : ["size"]
	, "cols"   : ["newprice", "food"]
	, "vals"   : lambda isize, inewprice, ifood: f'=(pi/e)*AVERAGE({minprice.ref(isize,ifood,sheetref=True)},{price.ref((isize),("Izza",ifood),sheetref=True,debug=True)})'
	})
})

x.writexls("example.xlsx")