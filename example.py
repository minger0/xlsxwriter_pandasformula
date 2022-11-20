import pandas as pd
import xlsxwriter_pandasformula as x

x.Constant('pi',3.14)
x.Constant('e',2.72)

price = x.View({
	"sheet"  : "price example"
, "anchor" : [0, 0]
, "value"  : pd.read_csv("exampleinput.csv", index_col=[0], header=[0,1])
}) # (size), (vendor, food)

minprice = x.View({
	"sheet"  : "price example"
, "anchor" : [1, 7]
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
, "value"  : x.Formula({
		"avgprice" : ['average']
	, "vendor" : x.dom['vendor']
	, "rows"   : ["avgprice"]
	, "cols"   : ["vendor"]
	, "vals"   : lambda iavgprice, ivendor: f'=MIN({price.ref((x.ALL),(ivendor,x.ALL))})'
	})
})

pricenewPizz = x.View({
	"sheet"  : "price example"
, "anchor" : [0, 16]
, "value"  : x.Formula({
		"size"   : x.dom['size']
	, "newprice": ["new price Pizz"]
	, "food"   : x.dom["food"]
	, "rows"   : ["size"]
	, "cols"   : ["newprice", "food"]
	, "vals"   : lambda isize, inewprice, ifood: f'=(pi/e)*AVERAGE({minprice.ref(isize,ifood)},{price.ref((isize),("Izza",ifood))})'
	})
})


x.writexls("example.xlsx")

