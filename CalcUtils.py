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

def GetStandardError(datasetLength, std):
    return std / math.sqrt(datasetLength)

def GetStandardErrorDifference(firstDatasetSE, firstDatasetLength, secondDatasetSE, secondDatasetLength):
    return math.sqrt((math.pow(firstDatasetSE, 2) / firstDatasetLength) + (math.pow(secondDatasetSE, 2) / secondDatasetLength))

def GetLSD(tAsteriskVal, SEDiff):
    return tAsteriskVal * SEDiff

def GetTStatistic(firstDatasetMean, secondDatasetMean, SEDiff):
    return (firstDatasetMean - secondDatasetMean) / SEDiff

def IsCriticalDifference(tStatistic, tAsteriskValue):
    return math.fabs(tStatistic) > tAsteriskValue

def StudentsT(firstDatasetMean, secondDatasetMean, LSD):
    return (firstDatasetMean - secondDatasetMean) / LSD


DS = [12, 8, 6, 6]
DS2 = [2, 1, 2, 5]


print(GetLowerAndUpperCl(DS2)[0] - GetLowerAndUpperCl(DS)[0])
print(GetLowerAndUpperCl(DS2)[1] - GetLowerAndUpperCl(DS)[1])