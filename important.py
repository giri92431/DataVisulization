
# df = pd.ExcelFile("Beer.xlsx" )
	


# # for i in df.sheet_names:
# # 	print(i.replace(" ","")+" = "+" df.parse('"+ i +"')")


# MarketSizes =  df.parse('Market Sizes')
# MarketSizes_002 =  df.parse('Market Sizes (002)')
# MarketSizes_003 =  df.parse('Market Sizes (003)')
# MarketSizes_004 =  df.parse('Market Sizes (004)')
# MarketSizes_005 =  df.parse('Market Sizes (005)')

# MarketSizes.to_pickle("MarketSizes.pickle")
# MarketSizes_002.to_pickle("MarketSizes_002.pickle")
# MarketSizes_003.to_pickle("MarketSizes_003.pickle")
# MarketSizes_004.to_pickle("MarketSizes_004.pickle")
# MarketSizes_005.to_pickle("MarketSizes_005.pickle")


# df1 = pd.read_pickle("MarketSizes.pickle")
# df2 = pd.read_pickle("MarketSizes_002.pickle")
# df3 = pd.read_pickle("MarketSizes_003.pickle")
# df4 = pd.read_pickle("MarketSizes_004.pickle")
# df5 = pd.read_pickle("MarketSizes_005.pickle")

# frames = [df1,df2,df3,df4,df5]

# result = pd.concat(frames)

# result.to_csv("beer.csv")
