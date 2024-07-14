import math

import numpy as np
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


def get_standard_error_diff(setsStandardError, minimumSampleSize):
    return np.sqrt(2 * setsStandardError / minimumSampleSize)


# Function to get the t-statistic and check if it's a critical difference
def get_t_statistic(t_critical_value, setsStandardError, minimumSampleSize, firstTreatmentMean, secondTreatmentMean):
    sed = get_standard_error_diff(setsStandardError, minimumSampleSize)
    t_stat = abs(firstTreatmentMean - secondTreatmentMean) / sed
    is_critical = t_stat > t_critical_value
    return t_stat, is_critical


def calculate_significant_letters(SortedTreatementDictionary, t_critical_value, setsStandardError, minimumSampleSize):
    keys = list(SortedTreatementDictionary.keys())
    sigByKey = {key: set() for key in keys}
    sigByKey[keys[0]] = 'A'
    LetterCounter = 'A'
    calculate_signficant_letters_recursion(SortedTreatementDictionary, t_critical_value, setsStandardError,
                                           minimumSampleSize, sigByKey, keys,
                                           LetterCounter)
    return sigByKey


def calculate_signficant_letters_recursion(SortedTreatementDictionary, t_critical_value, setsStandardError,
                                           minimumSampleSize, sigByKey, keys,
                                           LetterCounter):
    def IncrementLetterCounter(LetterCounter):
        AsciiOfLetter = ord(LetterCounter)
        IncrementedLettersAscii = AsciiOfLetter + 1
        return chr(IncrementedLettersAscii)

    for index, key in enumerate(keys[1:]):
        _, is_critical_dif = get_t_statistic(t_critical_value, setsStandardError, minimumSampleSize,
                                             SortedTreatementDictionary[key], SortedTreatementDictionary[keys[0]])

        if is_critical_dif:
            LetterCounter = IncrementLetterCounter(LetterCounter)
            sigByKey[key].add(LetterCounter)
            for secondKey in keys[1:]:
                _, is_critical_dif_second = get_t_statistic(t_critical_value, setsStandardError, minimumSampleSize,
                                                            SortedTreatementDictionary[key],
                                                            SortedTreatementDictionary[secondKey])
                if not is_critical_dif_second:
                    sigByKey[secondKey].add(LetterCounter)

            calculate_signficant_letters_recursion(SortedTreatementDictionary, t_critical_value, setsStandardError,
                                                   minimumSampleSize, sigByKey,
                                                   keys[index + 1:], LetterCounter)
            return

        sigByKey[key].add(LetterCounter)


def GetStandardError(datasetLength, std):
    return std / math.sqrt(datasetLength)


def GetStandardErrorDifference(firstDatasetSE, firstDatasetLength, secondDatasetSE, secondDatasetLength):
    return math.sqrt(
        (math.pow(firstDatasetSE, 2) / firstDatasetLength) + (math.pow(secondDatasetSE, 2) / secondDatasetLength))


def GetLSD(t_critical, mse, sampleSize):
    return t_critical * np.sqrt(2 * mse / sampleSize)


def GetTStatistic(firstDatasetMean, secondDatasetMean, SEDiff):
    return math.fabs(firstDatasetMean - secondDatasetMean) / SEDiff


def IsCriticalDifference(tStatistic, tAsteriskValue):
    return math.fabs(tStatistic) > tAsteriskValue


def StudentsT(firstDatasetMean, secondDatasetMean, LSD):
    return math.fabs(firstDatasetMean - secondDatasetMean) - LSD



