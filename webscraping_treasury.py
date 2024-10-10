import requests, json
import pandas as pd, os, time
from bs4 import BeautifulSoup
from openpyxl import load_workbook


# Function to change the width of columns in an Excel sheet
def change_column_width():
    wb = load_workbook('output.xlsx')
    ws = wb['Sheet1']
    for column in ws.columns:
        ws.column_dimensions[column[0].column_letter].width = 24.5
    wb.save('output.xlsx')


# Function to retrieve JSON data from a URL
def get_json():
    # URL for getting the XML links from the website
    url = "https://www.treasurydirect.gov/TA_WS/securities/search?startDate=2022-01-01&endDate=2022-12-31&compact=true&dateFieldName=auctionDate&format=jsonp&pdfFilenameAnnouncement=notNull&callback=jQuery36006008301254248725_1693045025835&filterscount=0&groupscount=2&group0=z3a&group1=t3a&pagenum=0&pagesize=500&recordstartindex=0&recordendindex=20&_=1693045040337"
    
    # Send a GET request to the URL
    response = requests.get(url)
    
    # Extract the data from the response
    response_list = response.text.split("(")[1].split(")")[0]
    
    # Parse the data into a JSON format
    response_list = json.loads(response_list)

    # Initialize a list to store XML URLs
    all_xml_urls = []
    
    # Loop through the JSON data
    for r in response_list:
        # Check if the entry is a "Bill" or "Note"
        if r["c"] == "Bill" or r["c"] == "Note":
            # Build the XML URL and add it to the list
            all_xml_urls.append("https://www.treasurydirect.gov/xml/" + r["ia1"])

    # Print the number of XML files found
    print("Found {} XML files".format(len(all_xml_urls)))

    # Return the list of XML URLs
    return all_xml_urls


# Main function to process XML files and update Excel output
def main(all_xml_urls):
    # Loop through each XML URL
    for url in all_xml_urls:
        print("Scraping XML file: {}".format(url))
        
        # Retry mechanism in case of network errors
        while True:
            try:
                response = requests.get(url)
                break
            except Exception as e:
                print(f"Error, retrying...{e}")
                time.sleep(15)

        # Parse XML content with BeautifulSoup
        soup = BeautifulSoup(response.content, 'lxml-xml')

        # Initialize a list to store data dictionaries
        data = []

        # Loop through each 'td:AuctionData' tag in the XML
        for tag in soup.find_all('td:AuctionData'):
            # Extract data and add each tag to a dictionary and keep adding it to the data list
            data.append({
                'SecurityTermWeekYear': tag.SecurityTermWeekYear.text,
                'SecurityTermDayMonth': tag.SecurityTermDayMonth.text,
                'SecurityType': tag.SecurityType.text,
                'CUSIP' : tag.CUSIP.text,
                'AnnouncementDate': tag.AnnouncementDate.text,
                'AuctionDate': tag.AuctionDate.text,
                'IssueDate': tag.IssueDate.text,
                'MaturityDate': tag.MaturityDate.text,
                'OfferingAmount': tag.OfferingAmount.text,
                'CompetitiveTenderAccepted': tag.CompetitiveTenderAccepted.text,
                'NonCompetitiveTenderAccepted': tag.NonCompetitiveTenderAccepted.text,
                'TreasuryDirectTenderAccepted': tag.TreasuryDirectTenderAccepted.text,
                'TypeOfAuction': tag.TypeOfAuction.text,
                'CompetitiveClosingTime': tag.CompetitiveClosingTime.text,
                'NonCompetitiveClosingTime': tag.NonCompetitiveClosingTime.text,
                'NetLongPositionReport': tag.NetLongPositionReport.text,
                'MaxAward': tag.MaxAward.text,
                'MaxSingleBid': tag.MaxSingleBid.text,
                'CompetitiveBidDecimals': tag.CompetitiveBidDecimals.text,
                'CompetitiveBidIncrement': tag.CompetitiveBidIncrement.text,
                'AllocationPercentageDecimals' : tag.AllocationPercentageDecimals.text,
                'MinBidAmount': tag.MinBidAmount.text,
                'MultiplesToBid': tag.MultiplesToBid.text,
                'MinToIssue': tag.MinToIssue.text,
                'MultiplesToIssue': tag.MultiplesToIssue.text,
                'MatureSecurityAmount': tag.MatureSecurityAmount.text,
                'CurrentlyOutstanding': tag.CurrentlyOutstanding.text,
                'SOMAIncluded': tag.SOMAIncluded.text,
                'SOMAHoldings': tag.SOMAHoldings.text,
                'MaturingDate': tag.MaturingDate.text,
                'FIMAIncluded': tag.FIMAIncluded.text,
                'Series': tag.Series.text,
                'InterestRate': tag.InterestRate.text,
                'Spread': tag.Spread.text,
                'FirstInterestPaymentDate': tag.FirstInterestPaymentDate.text,
                'StandardInterestPayment': tag.StandardInterestPayment.text,
                'FrequencyInterestPayment': tag.FrequencyInterestPayment.text,
                'StrippableIndicator': tag.StrippableIndicator.text,
                'MinStripAmount': tag.MinStripAmount.text,
                'CorpusCUSIP': tag.CorpusCUSIP.text,
                'TINTCUSIP1': tag.TINTCUSIP1.text,
                'TINTCUSIP2': tag.TINTCUSIP2.text,
                'ReOpeningIndicator': tag.ReOpeningIndicator.text,
                'OriginalIssueDate': tag.OriginalIssueDate.text,
                'BackDated': tag.BackDated.text,
                'BackDatedDate': tag.BackDatedDate.text,
                'LongShortNormalCoupon': tag.LongShortNormalCoupon.text,
                'LongShortCouponFirstIntPmt': tag.LongShortCouponFirstIntPmt.text,
                'InflationIndexSecurity': tag.InflationIndexSecurity.text,
                'FloatingRate': tag.FloatingRate.text,
                'RefCPIIssueDate': tag.RefCPIIssueDate.text,
                'RefCPIDatedDate': tag.RefCPIDatedDate.text,
                'IndexRatioOnIssueDate': tag.IndexRatioOnIssueDate.text,
                'CPIBasePeriod': tag.CPIBasePeriod.text,
                'TIINConversionFactor': tag.TIINConversionFactor.text,
                'AccruedInterest': tag.AccruedInterest.text,
                'DatedDate': tag.DatedDate.text,
                'AnnouncedCUSIP': tag.AnnouncedCUSIP.text,
                'UnadjustedPrice': tag.UnadjustedPrice.text,
                'UnadjustedAccruedInterest': tag.UnadjustedAccruedInterest.text,
                'AnnouncementPDFName': tag.AnnouncementPDFName.text,
                'OriginalDatedDate': tag.OriginalDatedDate.text,
                'AdjustedAmountCurrentlyOutstanding': tag.AdjustedAmountCurrentlyOutstanding.text,
                'NLPExclusionAmount': tag.NLPExclusionAmount.text,
                'MaximumNonCompAward': tag.MaximumNonCompAward.text,
                'AdjustedAccruedInterest': tag.AdjustedAccruedInterest.text,
                'Callable': tag.Callable.text,
                'CallDate': tag.CallDate.text,
                'LongShortCouponFirstIntPmt': tag.LongShortCouponFirstIntPmt.text,
                'PrimaryDealerTendered': tag.PrimaryDealerTendered.text,
                'PrimaryDealerAccepted': tag.PrimaryDealerAccepted.text,
                'DirectBidderTendered': tag.DirectBidderTendered.text,
                'DirectBidderAccepted': tag.DirectBidderAccepted.text,
                'IndirectBidderTendered': tag.IndirectBidderTendered.text,
                'IndirectBidderAccepted': tag.IndirectBidderAccepted.text,
                'CompetitiveTendered': tag.CompetitiveTendered.text,
                'CompetitiveAccepted': tag.CompetitiveAccepted.text,
                'NonCompetitiveAccepted': tag.NonCompetitiveAccepted.text,
                'SOMATendered': tag.SOMATendered.text,
                'SOMAAccepted': tag.SOMAAccepted.text,
                'FIMATendered': tag.FIMATendered.text,
                'FIMAAccepted': tag.FIMAAccepted.text,
                'TotalTendered': tag.TotalTendered.text,
                'TotalAccepted': tag.TotalAccepted.text,
                'BidToCoverRatio': tag.BidToCoverRatio.text,
                'ReleaseTime': tag.ReleaseTime.text,
                'AmountAcceptedBelowLowRate': tag.AmountAcceptedBelowLowRate.text,
                'HighAllocationPercentage': tag.HighAllocationPercentage.text,
                'LowDiscountRate': tag.LowDiscountRate.text,
                'HighDiscountRate': tag.HighDiscountRate.text,
                'MedianDiscountRate': tag.MedianDiscountRate.text,
                'LowYield': tag.LowYield.text,
                'HighYield': tag.HighYield.text,
                'MedianYield': tag.MedianYield.text,
                'LowDiscountMargin': tag.LowDiscountMargin.text,
                'HighDiscountMargin': tag.HighDiscountMargin.text,
                'MedianDiscountMargin': tag.MedianDiscountMargin.text,
                'LowPrice': tag.LowPrice.text,
                'HighPrice': tag.HighPrice.text,
                'MedianPrice': tag.MedianPrice.text,
                'TIINConversionFactor': tag.TIINConversionFactor.text,
                'AccruedInterest': tag.AccruedInterest.text,
                'StandardInterestPayment': tag.StandardInterestPayment.text,
                'InterestRate': tag.InterestRate.text,
                'Spread': tag.Spread.text,
                'OriginalCUSIP': tag.OriginalCUSIP.text,
                'UnadjustedPrice': tag.UnadjustedPrice.text,
                'UnadjustedAccruedInterest': tag.UnadjustedAccruedInterest.text,
                'TreasuryDirectAccepted': tag.TreasuryDirectAccepted.text,
                'InvestmentRate': tag.InvestmentRate.text,
                'AdjustedPrice': tag.AdjustedPrice.text,
                'AdjustedAccruedInterest': tag.AdjustedAccruedInterest.text,
                'IndexRatio': tag.IndexRatio.text,
                'FRNIndexDeterminationDate': tag.FRNIndexDeterminationDate.text,
                'FRNIndexDeterminationRate': tag.FRNIndexDeterminationRate.text,
                'ResultsPDFName': tag.ResultsPDFName.text
            })

        # Convert data list into a DataFrame
        df = pd.DataFrame(data)

        # Append or create Excel output
        if os.path.exists("output.xlsx"):
            df1 = pd.read_excel("output.xlsx")
            df2 = pd.concat([df1, df], ignore_index=True)
            df2.to_excel("output.xlsx", index=False)
        else:
            df.to_excel("output.xlsx", index=False)

    # Finalize DataFrame and Excel output
    df = pd.DataFrame([])
    if os.path.exists("output.xlsx"):
        df1 = pd.read_excel("output.xlsx")
        df2 = pd.concat([df1, df], ignore_index=True)
        df2.to_excel("output.xlsx", index=False)
    else:
        df.to_excel("output.xlsx", index=False)


# This block of code is often used to define the entry point of the script.
# It checks if the script is being run as the main program (not imported as a module).
# If the script is the main program, the code inside this block will be executed.
if __name__ == "__main__":
    print("Program Started\n")
    
    # Get a list of XML URLs by calling the 'get_json' function
    all_xml_urls = get_json()
    
    # Process the XML URLs and update the Excel output using the 'main' function
    main(all_xml_urls)
    
    change_column_width()
    print("\nProgram Ended")
