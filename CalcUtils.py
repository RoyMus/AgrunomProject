import math
from scipy.stats import t

def GetMean(dataset):
    sum = 0
    for number in dataset:
        sum += number

    return sum / len(dataset)

def CalcStd(dataset, mean):
    deltaSumSquared = 0
    for number in dataset:
        deltaSumSquared += math.pow((number - mean), 2)

    return math.sqrt(deltaSumSquared / (len(dataset) - 1))

def FindTAsteriskValue(datasetsLength):
    return t.ppf(1 - 0.025, datasetsLength - 1)

def CalcMarginOfError(datasetsLength, std, tAsterisk):
    return tAsterisk * std / math.sqrt(datasetsLength)

def GetLowerAndUpperCl(dataset):
    mean = GetMean(dataset)
    std = CalcStd(dataset, mean)
    tAsteriskVal = FindTAsteriskValue(len(dataset))
    me = CalcMarginOfError(len(dataset), std, tAsteriskVal)
    return (mean - me, mean + me)