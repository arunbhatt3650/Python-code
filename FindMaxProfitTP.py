import xlrd

loc = r"C:\Arun\Trading\Test\BN OTM1 CE PE SL 50 PT ATC.xls"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


# For row 0 and column 0
# print(sheet.cell_value(0, 0))
# print(sheet.nrows)
# print(sheet.ncols)


def findprofit(targetProfit):
    pnl = 0
    mdd = 0
    dd = 0
    adjustedPnl = 0
    for i in range(1, sheet.nrows):
        if abs(float(sheet.cell_value(i, 3)) / 25) > targetProfit:
            adjustedPnl = targetProfit * 25
        else:
            adjustedPnl = float(sheet.cell_value(i, 2))
        pnl += adjustedPnl
        dd += adjustedPnl
        if dd >= 0:
            dd = 0
        if dd < mdd:
            mdd = dd
    print(targetProfit, mdd, pnl)
    return pnl


#    print(pnl)
maxProfit = 0
opttp = 0
for ind in range(40, 200):
    profit = findprofit(ind)
    #    print(ind, profit)
    if profit > maxProfit:
        maxProfit = profit
        opttp = ind
print("maxprofit", maxProfit, "optTP", opttp)
