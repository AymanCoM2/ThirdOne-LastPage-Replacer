from b_headerFooterData import runHeaderFooterQuery, runRowQuery
from decimal import Decimal


def headerFooterParsing(docNumber):
    headerFooterResult = runHeaderFooterQuery(docNumber)
    hfDict = {}
    for item in headerFooterResult:
        _, description, company, code, _, start_date, end_date, primary_id, * \
            decimals, status, _ = item
        decimals = [float(d) for d in decimals]

        start_date = start_date.strftime('%Y/%m/%d')
        end_date = end_date.strftime('%Y/%m/%d')
        hfDict = {
            'Description': description,
            'Company': company,
            'Code': code,
            'StartDate': start_date,
            'EndDate': end_date,
            'PrimaryID': primary_id,
            'Decimals': decimals,
            'Status': status
        }

    headerList = [
        "", "",
        "", hfDict['PrimaryID'],
        hfDict['EndDate'], hfDict['StartDate'],
        "", hfDict['Description'],
        hfDict['Code'], hfDict['Company'],
    ]

    footerList = [str(hfDict['Decimals'][0]),
                  str(hfDict['Decimals'][2])+"  " +
                  str(hfDict['Decimals'][1]) + "%",
                  str(hfDict['Decimals'][3]),
                  str(hfDict['Decimals'][4]),
                  str(hfDict['Decimals'][5])
                  ]

    return headerList, footerList


def convert_to_rounded_string(value):
    if isinstance(value, Decimal):
        return str(round(float(value), 1))
    else:
        return str(value)


def rowsParsing(docNumber):
    counter = 1
    rowResult = runRowQuery(docNumber)
    finalDataList = []
    for rRow in rowResult:
        internalList = [
            convert_to_rounded_string(rRow[10]),
            convert_to_rounded_string(rRow[9]),
            convert_to_rounded_string(rRow[8]) + "%",
            convert_to_rounded_string(rRow[7]),
            convert_to_rounded_string(
                rRow[5]) + " " + convert_to_rounded_string(rRow[6]),
            convert_to_rounded_string(rRow[4]),
            convert_to_rounded_string(rRow[3]),
            convert_to_rounded_string(rRow[2]),
            convert_to_rounded_string(counter)
        ]
        finalDataList.append(internalList)
        counter = counter + 1
    return finalDataList


