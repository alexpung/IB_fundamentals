# IB_fundamentals
A short example shows how to dump Japanese stocks fundamentals with IB API from a stock list 
from excel. Could be modified to accept a list of tickers

Note that There is a limitation of 60 reqFundamentalData requests that can be made in a 10 minute period.
So an async implementation is good enough for the purpose of demonstration.