import pandas as pd
import numpy as np 
from threading import Thread


df = pd.read_pickle("beer.pickle")

print (df["Subcategory"].value_counts())
sdfsdfsdf