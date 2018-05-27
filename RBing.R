#
# configure dates, i.e. last 3 full days
startDate <- as.Date(Sys.Date() - 3)
endDate <- as.Date(Sys.Date() - 1)

# set working directory
setwd("O:/Path/To/Your/Project/Folder")

# configure list of BingAds clients
config <- data.frame(CustomerAccountId = c(""), 
                     CustomerId = c(""), 
                     Password = c(""), 
                     UserName = c(""), 
                     Format = c("Csv"), 
                     DeveloperToken = c(""), 
                     stringsAsFactors = FALSE)

submitGenerateReport <- function (CustomerAccountId, CustomerId, DeveloperToken, Password, UserName, Format, from, to) {
  
  if (!require("httr")) install.packages("httr")
  stopifnot(library(httr, logical.return = TRUE))
  if (!require("XML")) install.packages("XML")
  stopifnot(library(XML, logical.return = TRUE))
  
  body <- "<s:Envelope xmlns:s='http://schemas.xmlsoap.org/soap/envelope/'><s:Header><h:ApplicationToken i:nil='true' xmlns:h='https://bingads.microsoft.com/Reporting/v11' xmlns:i='http://www.w3.org/2001/XMLSchema-instance'/><h:AuthenticationToken i:nil='true' xmlns:h='https://bingads.microsoft.com/Reporting/v11' xmlns:i='http://www.w3.org/2001/XMLSchema-instance'/>"
  body <- paste0(body, "<h:CustomerAccountId xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", CustomerAccountId, "</h:CustomerAccountId>")
  body <- paste0(body, "<h:CustomerId xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", CustomerId, "</h:CustomerId>")
  body <- paste0(body, "<h:DeveloperToken xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", DeveloperToken, "</h:DeveloperToken>")
  body <- paste0(body, "<h:Password xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", Password, "</h:Password>")
  body <- paste0(body, "<h:UserName xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", UserName, "</h:UserName>")
  body <- paste0(body, "</s:Header><s:Body><SubmitGenerateReportRequest xmlns='https://bingads.microsoft.com/Reporting/v11'><ReportRequest i:type='DestinationUrlPerformanceReportRequest' xmlns:i='http://www.w3.org/2001/XMLSchema-instance'><Format>", Format, "</Format><Language i:nil='true'/><ReportName>My Destination Url Report</ReportName><ReturnOnlyCompleteData i:nil='true'/><Aggregation>Daily</Aggregation>")
  body <- paste0(body, "<Columns><DestinationUrlPerformanceReportColumn>TimePeriod</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>AccountId</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>AccountName</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>CampaignName</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>DestinationUrl</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>FinalURL</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>TrackingTemplate</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>Impressions</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>Clicks</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>Spend</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>AccountNumber</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>CampaignId</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>AdGroupName</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>AdGroupId</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>AdId</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>CurrencyCode</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>AdDistribution</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>Ctr</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>AverageCpc</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>AveragePosition</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>Conversions</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>ConversionRate</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>CostPerConversion</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>DeviceType</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>Language</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>BidMatchType</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>DeliveredMatchType</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>Network</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>TopVsOther</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>DeviceOS</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>Assists</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>Revenue</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>ReturnOnAdSpend</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>CostPerAssist</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>RevenuePerConversion</DestinationUrlPerformanceReportColumn><DestinationUrlPerformanceReportColumn>RevenuePerAssist</DestinationUrlPerformanceReportColumn></Columns>", "<Filter><AdDistribution i:nil='true'/><LanguageCode i:nil='true' xmlns:a='http://schemas.microsoft.com/2003/10/Serialization/Arrays'/></Filter>","<Scope><AccountIds xmlns:a='http://schemas.microsoft.com/2003/10/Serialization/Arrays'><a:long>", CustomerAccountId, "</a:long></AccountIds><AdGroups i:nil='true'/><Campaigns i:nil='true'/></Scope>")
  body <- paste0(body, "<Time><CustomDateRangeEnd><Day>", format(to, "%d"), "</Day><Month>", format(to, "%m"), "</Month><Year>", format(to, "%Y"), "</Year></CustomDateRangeEnd><CustomDateRangeStart><Day>", format(from, "%d"), "</Day><Month>", format(from, "%m"), "</Month><Year>", format(from, "%Y"), "</Year></CustomDateRangeStart><PredefinedTime i:nil='true'/><!-- <PredefinedTime>Yesterday</PredefinedTime> --></Time></ReportRequest></SubmitGenerateReportRequest></s:Body></s:Envelope>")
  body <- gsub(pattern = '\\\n', replacement = "", x = body)
  
  #
  # Windows users need to get a file with certificates
  if(!file.exists("cacert.pem") & url.exists("http://curl.haxx.se/ca/cacert.pem")){
    download.file(url = "http://curl.haxx.se/ca/cacert.pem", destfile = "cacert.pem")
  }
  
  API_URL <- "https://reporting.api.bingads.microsoft.com/Api/Advertiser/Reporting/v11/ReportingService.svc"
  
  # Force TLS1.2 as per Microsoft's requirements 
  # https://blogs.msdn.microsoft.com/bing_ads_api/2018/02/02/mandatory-upgrade-required-to-tls1-2/
  set_config(config(sslversion = 6))
  r <- POST(url = API_URL, 
            add_headers(SOAPAction = "SubmitGenerateReport", 
                        'Content-Type' = "text/xml; charset=utf-8"), 
            body = body)
  data <- xmlParse(content(r, "text"))
  ReportRequestId <- xmlToList(data)$Body$SubmitGenerateReportResponse$ReportRequestId 
  
  return(ReportRequestId)
}
pollGenerateReport <- function (CustomerAccountId, CustomerId, DeveloperToken, Password, UserName, Format, ReportRequestId) {
  
  if (!require("httr")) install.packages("httr")
  stopifnot(library(httr, logical.return = TRUE))
  if (!require("XML")) install.packages("XML")
  stopifnot(library(XML, logical.return = TRUE))
  if (!require("stringr")) install.packages("stringr")
  stopifnot(library(stringr, logical.return = TRUE))
  #
  # dynamic PollGenerateReport
  body <- "<s:Envelope xmlns:s='http://schemas.xmlsoap.org/soap/envelope/'><s:Header><h:ApplicationToken i:nil='true' xmlns:h='https://bingads.microsoft.com/Reporting/v11' xmlns:i='http://www.w3.org/2001/XMLSchema-instance'/><h:AuthenticationToken i:nil='true' xmlns:h='https://bingads.microsoft.com/Reporting/v11' xmlns:i='http://www.w3.org/2001/XMLSchema-instance'/>"
  body <- paste0(body, "<h:CustomerAccountId xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", CustomerAccountId, "</h:CustomerAccountId>")
  body <- paste0(body, "<h:CustomerId xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", CustomerId, "</h:CustomerId>")
  body <- paste0(body, "<h:DeveloperToken xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", DeveloperToken, "</h:DeveloperToken>")
  body <- paste0(body, "<h:Password xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", Password, "</h:Password>")
  body <- paste0(body, "<h:UserName xmlns:h='https://bingads.microsoft.com/Reporting/v11'>", UserName, "</h:UserName>")
  body <- paste0(body, "</s:Header><s:Body><PollGenerateReportRequest xmlns='https://bingads.microsoft.com/Reporting/v11'><ReportRequestId>", ReportRequestId, "</ReportRequestId></PollGenerateReportRequest></s:Body></s:Envelope>")
  body <- gsub(pattern = '\\\n', replacement = "", x = body)
  
  API_URL <- "https://reporting.api.bingads.microsoft.com/Api/Advertiser/Reporting/v11/ReportingService.svc"
  
  #
  # loop until the report is ready for download
  ReportRequestStatus <- ""
  while(ReportRequestStatus != "Success") {

    # Force TLS1.2 as per Microsoft's requirements 
    # https://blogs.msdn.microsoft.com/bing_ads_api/2018/02/02/mandatory-upgrade-required-to-tls1-2/
    set_config(config(sslversion = 6))
    r <- POST(url = API_URL, 
              add_headers(SOAPAction = "PollGenerateReport", 
                          'Content-Type' = "text/xml; charset=utf-8"), 
              body = body)
    data <- xmlParse(content(r, "text"))
    ReportRequestStatus <- xmlToList(data)$Body$PollGenerateReportResponse$ReportRequestStatus$Status 
    
    Sys.sleep(5)
  }
  #
  # download, extract and load report
  if (ReportRequestStatus == "Success") {
    ReportDownloadUrl <- xmlToList(data)$Body$PollGenerateReportResponse$ReportRequestStatus$ReportDownloadUrl
    
    # Force TLS1.2 as per Microsoft's requirements 
    # https://blogs.msdn.microsoft.com/bing_ads_api/2018/02/02/mandatory-upgrade-required-to-tls1-2/
    set_config(config(sslversion = 6))
    r <- GET(url = ReportDownloadUrl, 
             write_disk(path = paste0("bing-", ReportRequestId, ".zip"), 
                        overwrite = TRUE))
    
    unzip(paste0("bing-", ReportRequestId, ".zip"))
    unlink(paste0("bing-", ReportRequestId, ".zip"))
    
    file.name <- paste0("bing-", ReportRequestId, ".", tolower(Format))
    file.rename(paste0(ReportRequestId, ".", tolower(Format)), file.name)
    
    df <- read.csv(file.name, nrows = 25, blank.lines.skip = FALSE)
    skiprows <- which(df == "")
    df <- read.csv(file.name, header = TRUE, fill = TRUE, skip = skiprows)
    unlink(file.name)    
  }
  return(df)
}

# loop over list of BingAds clients and download DestinationUrlPerformanceReportRequest report
report <- NULL
for(i in 1:nrow(config)) { 
	print(paste0("Fetching ", config$UserName[i], "... "))
	ReportRequestId <- submitGenerateReport(config$CustomerAccountId[i], config$CustomerId[i], config$DeveloperToken[i], config$Password[i], config$UserName[i], config$Format[i], startDate, endDate)
	df <- pollGenerateReport(config$CustomerAccountId[i], config$CustomerId[i], config$DeveloperToken[i], config$Password[i], config$UserName[i], config$Format[i], ReportRequestId)
	if(!is.null(df)) {
		print(paste0(nrow(df), " rows fetched. "))
		# remove footer row
		df <- df[!str_detect(df$GregorianDate, "microsoft") > 0, ]
		df$GregorianDate <- as.Date(df$GregorianDate)
		df$Ctr <- gsub(pattern = "%$", replacement = "", x = df$Ctr, ignore.case = TRUE)
		df$ConversionRate <- gsub(pattern = "%$", replacement = "", x = df$ConversionRate, ignore.case = TRUE)
		df$ReturnOnAdSpend <- gsub(pattern = "%$", replacement = "", x = df$ReturnOnAdSpend, ignore.case = TRUE)
		if(is.null(df$AverageCpm)) { 
		  df$AverageCpm <- "" 
		}
		if(is.null(df$PricingModel)) { 
		  df$PricingModel <- "" 
		}
		if(is.null(df$ExtendedCost)) { 
		  df$ExtendedCost <- ""
		}
		report <- rbind.fill(report, df)
	} else {
		print("Warning! Empty report.")
	}
	rm(ReportRequestId, df)
}

# Store report in CSV
write.csv(x = report, file = "Bing.csv", row.names = FALSE)