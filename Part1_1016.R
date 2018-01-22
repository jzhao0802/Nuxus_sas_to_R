#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#   Project: NEXXUS
#   Purpose: Loading and Cleaning data for Propensity Score Matching
#   Version: 0.1
#   Programmer: Xin Huang
#   Date: 09/09/2015
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# set up the referrence and load some required library and functions
path <- paste("\\\\plyvnas01\\statservices2\\CustomStudies\\Promotion Management\\",
              "2014\\NOVOLOG\\04 Codes\\Nexxus\\R_Version\\1016", 
              sep = "")
setwd(path)
source("match.R")

#install.packages("Reshape2")
require(ggplot2)
require(dplyr) # merge, very efficient
require(reshape2) #transponse
require(stringr) # pad zero in ids

prog = NULL
min_postwks = 4

# read in the sas data
# here chk.csv is the the csv version of sas data ims_rx_data_sample
ddHeader <- tolower(scan("for_sas_test_0929.csv", what = character(),
                         sep = ',', nlines = 1))
cc <- c(imsdr7 = "character",
        zipcode = "character",
        spec = "character",
        specialty = "character",
        product = "character",
        datadate = "character",
        week = "numeric",
        trx = "numeric",
        nrx = "numeric")

ims_rx_data_sample1 <- read.csv("for_sas_test_0929.csv", 
                                col.names = ddHeader,
                                colClasses=cc,
                                stringsAsFactors = FALSE)

# main function
part1_pre_process <- function(raw_xpo_data, 
                              campaign_prd = "VICTOZA",
                              pre_wks = 26, 
                              campaign_data = "campaign_139.csv",
                              study_rx = "nrx",
                              total_wks = 128) {
  #@parameters:
  #  Exponent data: Weekly data with Week #s
  #  Campaign_PRD:  Campaign product, all other products are considered as "Other"
  #  Pre_wks:       # of weeks to considedr for pre period
  #  Campaign_dat:  Campaign Contact HCP list excel file
  #  Study_RX:      Main match and analysis Rx variable
  #  Total_wks: 	  Total # of weeks of availabel data
  
  #@returns:
  #  matched data
  
#   raw_xpo_data = "ims_rx_data_sample1"
#   campaign_prd = "VICTOZA"
#   pre_wks = 26
#   campaign_data = "campaign_139"
#   study_rx = "nrx"
#   total_wks = 128
  
  # check the input
  if (!is.character(raw_xpo_data)) {
    stop("It should be the name of the XPO data, a character pointing to the data,
         not the data itself")
  }
  if (!all(match(c("imsdr7", "zipcode", "spec", "specialty", "product",
                   "datadate", "week", "trx", "nrx"),
                 colnames(get(raw_xpo_data)), nomatch = FALSE))) {
    stop("XPO data should contains all the following columns: 
         == imsdr7, zipcode, spec, specialty, product,
            datadate, week, trx, nrx ==")
  }
  if (!is.character(campaign_data)) {
    stop("It should be the name of the Campaign data, a character pointing to the data,
         not the data itself")
  }
  if (!tolower(study_rx) %in% c("nrx", "trx")) {
    stop("study_rx should just be NRX and TRX")
  }
  
  infile = "raw_xpo_data"
  min_postwks <- 4
  
  raw_data <- subset(get(get(infile)), as.numeric(imsdr7) < 9000000)
  raw_data[, "brand"] <- toupper(raw_data[, "product"])
  level <- c("imsdr7", "spec", "zipcode", "brand")
  
  # remove na function
  replace_na <- function(x) {
    if (is.numeric(x)) {
      x[is.na(x)] <- 0
    } else {
      x
    }
    x
  }
  
  raw_data[] <- lapply(raw_data, replace_na)

  aggregated <- raw_data %>%
    filter(as.numeric(imsdr7) < 9000000) %>%
    mutate(brand = toupper(product)) %>%
    group_by(imsdr7, spec, zipcode, brand, week) %>%
    summarize(trx = sum(trx), nrx = sum(nrx))

  # transponse at level imsdr7 + spec + zipcode + brand
  trx_t <- dcast(aggregated, imsdr7 + spec + zipcode + brand ~ week, 
                 value.var = "trx")
  colnames(trx_t)[(length(level) + 1) : ncol(trx_t)] <-
    paste0("trx_wk_", colnames(trx_t)[(length(level) + 1) : ncol(trx_t)])
  nrx_t <- dcast(aggregated, imsdr7 + spec + zipcode + brand ~ week, 
                 value.var = "nrx")
  colnames(nrx_t)[(length(level) + 1) : ncol(nrx_t)] <-
    paste0("nrx_wk_", colnames(nrx_t)[(length(level) + 1) : ncol(nrx_t)])
  
  # read in the spec information
  spec_grp <- read.csv("spec_grp.csv", header = TRUE, 
                       stringsAsFactors = FALSE)
  colnames(spec_grp) <- tolower(colnames(spec_grp))
  colnames(spec_grp)[1] <- "spec"
  
  # merge TRX and NRX together
  transposed <- full_join(trx_t, nrx_t, by = level)
  transposed[, "spec"][is.na(transposed[, "spec"])] <- "OTH"
  transposed <- left_join(transposed, spec_grp[, -3], by = "spec")
  transposed[, "spec_grp"] <- ifelse(is.na(transposed[, "grp_id"]), 1, 
                                     transposed[, "grp_id"])
  
  # read in zip region information
  ddHeader1 <- tolower(scan("ZIP_region.csv", what=character(),
                            sep=',', nlines=1))
  if (!all(match(c("zipcode", "state", "region", "lat_nbr",
                   "lon_nbr", "x_coord_albers_nbr",
                   "y_coord_albers_nbr"), ddHeader1, nomatch = FALSE))) {
    stop("zip-region data should contains all the following columns: 
         == zipcode, state, region, lat_nbr, lon_nbr, x_coord_albers_nbr,
            y_coord_albers_nbr ==")
  }
  cc1 <- c(zipcode = "character",
           state = "character",
           region = "character",
           lat_nbr = "numeric",
           lon_nbr = "numeric",
           x_coord_albers_nbr = "numeric",
           y_coord_albers_nbr = "numeric")
  zipregion <- read.csv("ZIP_region.csv", col.names = ddHeader1,
                        colClasses = cc1,
                        stringsAsFactors = FALSE)
  assign(gsub(" ", "", paste(get(infile), "_", "rx")),
         left_join(transposed, zipregion, by = "zipcode"))
  tmp <- get(gsub(" ", "", paste(get(infile), "_", "rx"))) 
  #1169 without coord
  
  tmp[, "brand"] <-ifelse(tmp[, "brand"] == campaign_prd, tmp[, "brand"], "OTHER")
  # replace the na in coords with 0
  tmp[, c("lat_nbr", "lon_nbr", "x_coord_albers_nbr","y_coord_albers_nbr")] <-
    lapply(tmp[, c("lat_nbr", "lon_nbr", "x_coord_albers_nbr","y_coord_albers_nbr")],
           function(x) ifelse(is.na(x), 0, x))
  
  
  # do the summary at level imsdr7 + spec + zipcode + brand
  tmp <- tmp[, setdiff(colnames(tmp), c("grp_id", "state"))]
  
  tmp[] <- lapply(tmp, replace_na)
  assign(gsub(" ", "",paste("ims_xpo_dat_", tolower(campaign_prd), "_rx")), 
         tmp %>% 
           group_by(imsdr7, spec, zipcode, brand, spec_grp, region, 
                      lat_nbr, lon_nbr, x_coord_albers_nbr, y_coord_albers_nbr) %>%
           summarise_each(funs(sum)))

  assign(gsub(" ", "",paste("ims_xpo_dat_", tolower(campaign_prd), "_rx")),
         as.data.frame(get(gsub(" ", "",
                                paste("ims_xpo_dat_",
                                      tolower(campaign_prd), 
                                      "_rx")))))
  
  # read in the campaigm data
  ddHeader2 <- tolower(scan(paste(campaign_data, ".csv", sep = ""),
                            what=character(), 
                            sep=',', nlines=1))
  if (!all(match(c("contactid", "engagement_date", "camapign_date",
                   "campaignid", "imsid"), ddHeader2, nomatch = FALSE))) {
    stop("campaign data should contains all the following columns: 
         == contactid, engagement_date, camapign_date, campaignid, imsid ==")
  }
  cc2 <- c(contactid = "character",
           engagement_date = "character",
           camapign_date = "character",
           campaignid = "character",
           imsid = "character")
  campaign_hcps <- read.csv(paste(campaign_data, ".csv", sep = ""),
                            col.names = ddHeader2,
                            colClasses = cc2,
                            stringsAsFactors = FALSE)  
  campaign_hcps[, "imsdr7"] <- str_pad(campaign_hcps[, "imsid"], 7, pad = "0")
  campaign_hcps <- subset(campaign_hcps, imsdr7 != "0000000")
  campaign_hcps_f <- table(campaign_hcps$imsdr7)
  campaign_hcps_imsdr7 <- data.frame(imsdr7 = names(campaign_hcps_f),
                                     count = campaign_hcps_f[drop = TRUE],
                                     stringsAsFactors = FALSE)
  # remove the duplicates
  campaign_hcps_imsdr7_m <- campaign_hcps_imsdr7[campaign_hcps_imsdr7[, "count"] > 1, ]
  campaign_hcps_imsdr7_m1 <- 
    lapply(campaign_hcps_imsdr7_m[, 1], 
           function(x){
             tmp <- campaign_hcps[campaign_hcps[, "imsdr7"] == x, ]
             tmp1 <- arrange(tmp, imsdr7, engagement_date)
             tmp1[1,]
             })
  campaign_hcps_imsdr7_m1 <- do.call("rbind", campaign_hcps_imsdr7_m1)
  campaign_hcps_m <- 
    campaign_hcps[!campaign_hcps[, "imsdr7"] %in%  campaign_hcps_imsdr7_m1[, "imsdr7"], ]
  campaign_hcps <- rbind(campaign_hcps_m, campaign_hcps_imsdr7_m1)
  prog <- unique(campaign_hcps[ ,"campaignid"]) 
  
  
  mby <- "weeek" # or month
  
  # set up the locale for the date malipulation
  lct <- Sys.getlocale("LC_TIME"); Sys.setlocale("LC_TIME", "C")
  ## Sys.setlocale("LC_TIME", lct)
  cstart <- min(as.Date(campaign_hcps[, "engagement_date"]))
  datadate <- as.Date(get(raw_xpo_data)[1, "datadate"])
  # start week and end week of pre-period
  start_wk <- total_wks + as.numeric(format(cstart, "%U")) + 
    (as.numeric(format(cstart, "%Y")) - as.numeric(format(datadate, "%Y"))) * 52 -
    as.numeric(format(datadate, "%U")) - 26
  end_wk <- total_wks + as.numeric(format(cstart, "%U")) + 
    (as.numeric(format(cstart, "%Y")) - as.numeric(format(datadate, "%Y"))) * 52 -
    as.numeric(format(datadate, "%U")) - 1
  
  # set up the time points
  imsdate2 <- datadate
  from <- start_wk
  to <- end_wk
  
  line1 <- length(unique(campaign_hcps[, "imsdr7"]))
  
  ims_rx_1rec_nov <- 
    subset(get(gsub(" ", "",paste("ims_xpo_dat_", tolower(campaign_prd), "_rx"))),
           brand == campaign_prd)
  ims_rx_1rec_oth <- 
    subset(get(gsub(" ", "",paste("ims_xpo_dat_", tolower(campaign_prd), "_rx"))),
           brand != campaign_prd)
  
  line2 <- length(unique(ims_rx_1rec_nov[, "imsdr7"]))
  
  # change the colnames of the OTH dataset
  #col_flag <- grep("^trx_|^nrx", colnames(ims_rx_1rec_oth))
  colnames(ims_rx_1rec_oth) <- gsub("_wk_", "_oth_", colnames(ims_rx_1rec_oth))
  
  idx_brand <- which(colnames(ims_rx_1rec_nov) == "brand")
  ims_rx_1rec_hcp <- full_join(ims_rx_1rec_nov[, -idx_brand], 
                               ims_rx_1rec_oth[, -idx_brand], 
                               by =c("imsdr7", "spec", "zipcode",
                                     "spec_grp", "region", "lat_nbr",
                                     "lon_nbr", "x_coord_albers_nbr",
                                     "y_coord_albers_nbr"))
  
  
  ims_rx_1rec_hcp[] <- lapply(ims_rx_1rec_hcp, replace_na)
  
  hcps_rx_all <- ims_rx_1rec_hcp
  hcps_rx_all[, "total_sum_nrx"] <-
    apply(hcps_rx_all[, grep("nrx_wk_", colnames(hcps_rx_all))], 
          1,
         sum,
         na.rm = TRUE)
  hcps_rx_all[, "total_sum"] <-
    apply(hcps_rx_all[, grep("trx_wk_", colnames(hcps_rx_all))], 
          1,
          sum,
          na.rm = TRUE)
  
  writers <- length(unique(subset(hcps_rx_all, total_sum > 0)[, "imsdr7"]))
  
  # define outlier function
  outlier <- function(data = hcps_rx_all){
    hcps_rx_all <- as.data.frame(data)
    mkt_cut99 <- quantile(hcps_rx_all[, "total_sum"], probs = 0.99)
    top_value <- subset(hcps_rx_all, total_sum >= mkt_cut99)
    top_value <- top_value[order(top_value[, "total_sum"], decreasing = TRUE), ]
    top_value[, "diff"] <- 
      lag(top_value[, "total_sum"], 1) - top_value[, "total_sum"]
    top_value[, "pct_chg"] <- top_value[, "diff"] / mkt_cut99
    large_jump <- subset(top_value, pct_chg >= 0.168) # can be 5% - 25%
    large_jump <- 
      large_jump[order(large_jump[, "total_sum"], decreasing = TRUE), ]
    mkt_cut <- large_jump[nrow(large_jump), "total_sum"]
    if (mkt_cut <= 0) mkt_cut <- 99999
    
    hcps_rx_all <- subset(hcps_rx_all, total_sum <= mkt_cut) 
  }
  hcps_rx_all <- outlier(hcps_rx_all)
  
  line3 <- length(unique(hcps_rx_all[, "imsdr7"]))
  
  test_hcp <- campaign_hcps[, c("engagement_date", "imsdr7")]
  test_hcp[, "engagement_date"] <- ifelse(as.Date(test_hcp[, "engagement_date"]) >
                                            as.Date(imsdate2), 
                                          imsdate2,
                                          test_hcp[, "engagement_date"])
  test_hcp[, "from_week"] <- total_wks + 
    as.numeric(format(as.Date(test_hcp[, "engagement_date"], "%Y/%m/%d"), "%U")) -
    as.numeric(format(imsdate2, "%U"))+ 
    (as.numeric(format(as.Date(test_hcp[, "engagement_date"], "%Y/%m/%d"), "%Y")) -
       as.numeric(format(imsdate2, "%Y"))) * 52 - pre_wks
  test_hcp[, "week"] <- total_wks + 
    as.numeric(format(as.Date(test_hcp[, "engagement_date"], "%Y/%m/%d"), "%U")) -
    as.numeric(format(imsdate2, "%U"))+ 
    (as.numeric(format(as.Date(test_hcp[, "engagement_date"], "%Y/%m/%d"), "%Y")) -
       as.numeric(format(imsdate2, "%Y"))) * 52 - 1
  test_hcp[, "eligible"] <- ifelse(test_hcp[, "week"] <= total_wks - min_postwks,
                                   1, NA)
  
  # get the week dataset
  weeks <- table(subset(test_hcp, eligible == 1)[, "week"])
  weeks <- as.numeric(names(weeks))
  weeks <- weeks[order(weeks)]
  wks_diff <- weeks - weeks[1]
  weeks <- cbind(week = weeks, wks_diff = wks_diff)
  month <- ifelse(ceiling(weeks[, "wks_diff"] / 4) == 0, 1,
                  ceiling(weeks[, "wks_diff"] / 4))
  weeks <- cbind(weeks, month = month)
  weeks <- as.data.frame(weeks[, c("week", "month")])
  
  # merge test_hcp with week
  test_hcp <- left_join(test_hcp, weeks, by = "week")
  
  all_wks <- weeks[, "week"]
  n_wks <- length(weeks[, "week"])
  
  all_mth <- weeks[, "month"]
  n_mth <- length(weeks[, "month"])
  
  # for test dataset
  all_test_hcps <- inner_join(hcps_rx_all, test_hcp, by = "imsdr7")
  all_test_hcps <- subset(all_test_hcps, eligible == 1)
  
  all_test_hcps[, "sr_elig"] <- ifelse(is.na(all_test_hcps[, "spec"]) & 
                                         (is.na(all_test_hcps[, "spec"]) |
                                         nchar(all_test_hcps[, "region"]) == 0), 
                                       0, 1)
  all_test_hcps[, "spec_num"] <- all_test_hcps[, "spec_grp"]
  all_test_hcps[, "region"] <- ifelse(is.na(all_test_hcps[, "region"]) |
                                        nchar(all_test_hcps[, "region"]) == 0, 
                                      "UNKNOWN", all_test_hcps[, "region"])
  all_test_hcps[, "reg_num"] <- 
    (all_test_hcps[, "region"] == "MIDWEST") * 1 +
    (all_test_hcps[, "region"] == "NORTHEAST") * 2 +
    (all_test_hcps[, "region"] == "SOUTH") * 3 +
    (all_test_hcps[, "region"] == "WEST") * 4
  
  # for control group
  control_hcps <- anti_join(hcps_rx_all, test_hcp, by = "imsdr7")
  
  control_hcps[, "sr_elig"] <- ifelse(is.na(control_hcps[, "spec"]) & 
                                         (is.na(control_hcps[, "region"]) |
                                            nchar(control_hcps[, "region"]) == 0), 
                                       0, 1)
  control_hcps[, "spec_num"] <- control_hcps[, "spec_grp"]
  control_hcps[, "region"] <- ifelse(is.na(control_hcps[, "region"]) |
                                       nchar(control_hcps[, "region"]) == 0, 
                                      "UNKNOWN", control_hcps[, "region"])
  control_hcps[, "reg_num"] <- 
    (control_hcps[, "region"] == "MIDWEST") * 1 +
    (control_hcps[, "region"] == "NORTHEAST") * 2 +
    (control_hcps[, "region"] == "SOUTH") * 3 +
    (control_hcps[, "region"] == "WEST") * 4
  
  # doctor number in control and test
  line1x <- length(unique(all_test_hcps[, "imsdr7"]))
  line2x <- length(unique(control_hcps[, "imsdr7"]))
  
  # for descriptive
  campaign_week <- total_wks + as.numeric(format(cstart, "%U")) -
    as.numeric(format(imsdate2, "%U"))+ 
    (as.numeric(format(cstart, "%Y")) - as.numeric(format(imsdate2, "%Y"))) * 52 - 1
  post_len <- total_wks - campaign_week
  
  descriptive_dat <- ims_rx_1rec_hcp
  pre_nov_trx <- paste("trx_wk_", seq(1, campaign_week), sep = "")
  pre_nov_nrx <- paste("nrx_wk_", seq(1, campaign_week), sep = "")
  
  post_nov_trx <- paste("trx_wk_", seq(campaign_week + 1, total_wks), sep = "")
  post_nov_nrx <- paste("nrx_wk_", seq(campaign_week + 1, total_wks), sep = "")
  
  pre_oth_trx <- paste("trx_oth_", seq(1, campaign_week), sep = "")
  pre_oth_nrx <- paste("nrx_oth_", seq(1, campaign_week), sep = "")
  
  post_oth_trx <- paste("trx_oth_", seq(campaign_week + 1, total_wks), sep = "")
  post_oth_nrx <- paste("nrx_oth_", seq(campaign_week + 1, total_wks), sep = "")
  
  descriptive_dat[, "pre_nall"] <- apply(descriptive_dat[, pre_nov_trx],
                                         1,
                                         sum,
                                         na.rm = TRUE)
  descriptive_dat[, "pre_nall_nrx"] <- apply(descriptive_dat[, pre_nov_nrx],
                                             1,
                                             sum,
                                             na.rm = TRUE)
  descriptive_dat[, "pre_oall"] <- apply(descriptive_dat[, pre_oth_trx],
                                         1,
                                         sum,
                                         na.rm = TRUE)
  descriptive_dat[, "pre_oall_nrx"] <- apply(descriptive_dat[, pre_oth_nrx],
                                             1,
                                             sum,
                                             na.rm = TRUE)
  
  descriptive_dat[, "post_nall"] <- apply(descriptive_dat[, post_nov_trx],
                                         1,
                                         sum,
                                         na.rm = TRUE)
  descriptive_dat[, "post_nall_nrx"] <- apply(descriptive_dat[, post_nov_nrx],
                                             1,
                                             sum,
                                             na.rm = TRUE)
  descriptive_dat[, "post_oall"] <- apply(descriptive_dat[, post_oth_trx],
                                         1,
                                         sum,
                                         na.rm = TRUE)
  descriptive_dat[, "post_oall_nrx"] <- apply(descriptive_dat[, post_oth_nrx],
                                             1,
                                             sum,
                                             na.rm = TRUE)
  lb_nov_trx <- paste("trx_wk_", 
                      seq(campaign_week - 26, campaign_week),
                      sep = "")
  lb_nov_nrx <- paste("nrx_wk_",
                      seq(campaign_week - 26, campaign_week),
                      sep = "")
  
  lb_oth_trx <- paste("trx_oth_", 
                      seq(campaign_week - 26, campaign_week),
                      sep = "")
  lb_oth_nrx <- paste("nrx_oth_",
                      seq(campaign_week - 26, campaign_week),
                      sep = "")
  
  descriptive_dat[, "pre_nov"] <- apply(descriptive_dat[, lb_nov_trx],
                                        1,
                                        sum,
                                        na.rm = TRUE)
  descriptive_dat[, "pre_nov_nrx"] <- apply(descriptive_dat[, lb_nov_nrx],
                                            1,
                                            sum,
                                            na.rm = TRUE)
  
  descriptive_dat[, "pre_oth"] <- apply(descriptive_dat[, lb_oth_trx],
                                        1,
                                        sum,
                                        na.rm = TRUE)
  descriptive_dat[, "pre_oth_nrx"] <- apply(descriptive_dat[, lb_oth_nrx],
                                            1,
                                            sum,
                                            na.rm = TRUE)
  
  descriptive_dat[, "post_nov"] <- apply(descriptive_dat[, post_nov_trx],
                                         1,
                                         sum,
                                         na.rm = TRUE)
  descriptive_dat[, "post_nov_nrx"] <- apply(descriptive_dat[, post_nov_nrx],
                                              1,
                                              sum,
                                              na.rm = TRUE)
  descriptive_dat[, "post_oth"] <- apply(descriptive_dat[, post_oth_trx],
                                          1,
                                          sum,
                                          na.rm = TRUE)
  descriptive_dat[, "post_oth_nrx"] <- apply(descriptive_dat[, post_oth_nrx],
                                              1,
                                              sum,
                                              na.rm = TRUE)
  
  # adjust the post RX for comparing
  descriptive_dat[, "post_nov"] <- 
    descriptive_dat[, "post_nov"] * pre_wks / post_len
  descriptive_dat[, "post_nov_nrx"] <- 
    descriptive_dat[, "post_nov_nrx"] * pre_wks / post_len
  descriptive_dat[, "post_oth"] <- 
    descriptive_dat[, "post_oth"] * pre_wks / post_len
  descriptive_dat[, "post_oth_nrx"] <-
    descriptive_dat[, "post_oth_nrx"] * pre_wks / post_len

  descriptive_dat[, "pre_mkt_trx"] <- 
    descriptive_dat[, "pre_nov"] + descriptive_dat[, "pre_oth"]
  descriptive_dat[, "post_mkt_trx"] <- 
    descriptive_dat[, "post_nov"] + descriptive_dat[, "post_oth"]
  descriptive_dat[, "pre_mkt_nrx"] <- 
    descriptive_dat[, "pre_nov_nrx"] + descriptive_dat[, "pre_oth_nrx"]
  descriptive_dat[, "post_mkt_nrx"] <- 
    descriptive_dat[, "post_nov_nrx"] + descriptive_dat[, "post_oth_nrx"]
  
  # calculate the RX writers for Nov, MKT and competitors
  descriptive_dat[, "pre_writer_x"] <-  descriptive_dat[, "pre_nall"] >= 1
  descriptive_dat[, "pre_writer_m"] <- 
    descriptive_dat[, "pre_nall"] >= 1 | descriptive_dat[, "pre_oall"] >= 1
  descriptive_dat[, "pre_writer_c"] <-  descriptive_dat[, "pre_oall"] >= 1
  
  descriptive_dat[, "post_writer_x"] <-  descriptive_dat[, "post_nall"] >= 1
  descriptive_dat[, "post_writer_m"] <- 
    descriptive_dat[, "post_nall"] >= 1 | descriptive_dat[, "post_oall"] >= 1
  descriptive_dat[, "post_writer_c"] <-  descriptive_dat[, "post_oall"] >= 1
  
  # get the descriptive statistics 
  vars <- colnames(descriptive_dat)[grep("^pre|^post", colnames(descriptive_dat))]
  descriptives <- apply(descriptive_dat[, vars], 2, sum, na.rm = TRUE)
  ndr <- length(unique(descriptive_dat[, "imsdr7"]))
  
  vars_desc <- c("pre_writer_x",  "post_writer_x",  "pre_nov", "post_nov", 
                 "pre_nov_nrx", "post_nov_nrx",
                 "pre_writer_m",  "post_writer_m",  "pre_mkt_trx", "post_mkt_trx",
                 "pre_mkt_nrx", "post_mkt_nrx",
                 "pre_writer_c",  "post_writer_c",  "pre_oth", "post_oth",
                 "pre_oth_nrx", "post_oth_nrx")
  
  descriptives <- matrix(descriptives[vars_desc], ncol = 6, byrow = TRUE)
  descriptives <- as.data.frame(descriptives)
  colnames(descriptives) <- c("pre_writer", "post_writer", "pre_trx", "post_trx",
                              "pre_nrx", "post_nrx")
  
  descriptives <- cbind(Product = c("Product X", "Market", "Competitors"),
                        total_campaign_hcps = rep(ndr, 3), descriptives)

  # pre-period summary
  vars_diff <- setdiff(colnames(all_test_hcps), colnames(control_hcps))
  all_test_hcps_1 <- all_test_hcps[, setdiff(colnames(all_test_hcps),
                                             vars_diff)]
  all_test_hcps_1 <- cbind(all_test_hcps_1, group = "TEST")
  control_hcps_1 <- cbind(control_hcps, group = "CONTROL")
  for_pre_match_comp <- rbind(all_test_hcps_1, control_hcps_1)
  cut_week <- total_wks + as.numeric(format(cstart, "%U")) -
    as.numeric(format(imsdate2, "%U"))+ 
    (as.numeric(format(cstart, "%Y")) - 
       as.numeric(format(imsdate2, "%Y"))) * 52 - 1
  
  vars_period_3 <- seq(cut_week - 26, cut_week - 14)
  vars_period_2 <- seq(cut_week - 13, cut_week - 5)
  vars_period_1 <- seq(cut_week - 4, cut_week - 1)

  var_rel_trx_3 <- paste("trx_wk_", vars_period_3, sep = "")
  var_rel_trx_2 <- paste("trx_wk_", vars_period_2, sep = "")
  var_rel_trx_1 <- paste("trx_wk_", vars_period_1, sep = "")

  var_rel_trx_oth_3 <- paste("trx_oth_", vars_period_3, sep = "")
  var_rel_trx_oth_2 <- paste("trx_oth_", vars_period_2, sep = "")
  var_rel_trx_oth_1 <- paste("trx_oth_", vars_period_1, sep = "")

  var_rel_nrx_3 <- paste("nrx_wk_", vars_period_3, sep = "")
  var_rel_nrx_2 <- paste("nrx_wk_", vars_period_2, sep = "")
  var_rel_nrx_1 <- paste("nrx_wk_", vars_period_1, sep = "")
  
  var_rel_nrx_oth_3 <- paste("nrx_oth_", vars_period_3, sep = "")
  var_rel_nrx_oth_2 <- paste("nrx_oth_", vars_period_2, sep = "")
  var_rel_nrx_oth_1 <- paste("nrx_oth_", vars_period_1, sep = "")
  
  # trx
  for_pre_match_comp[, "rel_trx_3"] <- 
    apply(for_pre_match_comp[, var_rel_trx_3], 1, sum)
  for_pre_match_comp[, "rel_trx_2"] <- 
    apply(for_pre_match_comp[, var_rel_trx_2], 1, sum)
  for_pre_match_comp[, "rel_trx_1"] <-
    apply(for_pre_match_comp[, var_rel_trx_1], 1, sum)

  for_pre_match_comp[, "rel_trx_oth_3"] <- 
    apply(for_pre_match_comp[, var_rel_trx_oth_3], 1, sum)
  for_pre_match_comp[, "rel_trx_oth_2"] <- 
    apply(for_pre_match_comp[, var_rel_trx_oth_2], 1, sum)
  for_pre_match_comp[, "rel_trx_oth_1"] <-
    apply(for_pre_match_comp[, var_rel_trx_oth_1], 1, sum)
  
  #nrx
  for_pre_match_comp[, "rel_nrx_3"] <- 
    apply(for_pre_match_comp[, var_rel_nrx_3], 1, sum)
  for_pre_match_comp[, "rel_nrx_2"] <- 
    apply(for_pre_match_comp[, var_rel_nrx_2], 1, sum)
  for_pre_match_comp[, "rel_nrx_1"] <-
    apply(for_pre_match_comp[, var_rel_nrx_1], 1, sum)
  
  for_pre_match_comp[, "rel_nrx_oth_3"] <- 
    apply(for_pre_match_comp[, var_rel_nrx_oth_3], 1, sum)
  for_pre_match_comp[, "rel_nrx_oth_2"] <- 
    apply(for_pre_match_comp[, var_rel_nrx_oth_2], 1, sum)
  for_pre_match_comp[, "rel_nrx_oth_1"] <-
    apply(for_pre_match_comp[, var_rel_nrx_oth_1], 1, sum)

  for_pre_match_comp[, "specialty"] <- rep("ALL SPEC", nrow(for_pre_match_comp))

  
  ## by region
  #mean
  pre_monthly_data_mean <- 
  aggregate(cbind(rel_trx_3, rel_trx_2, rel_trx_1,
                  rel_nrx_3, rel_nrx_2, rel_nrx_1) ~ specialty + region +
              group, data = for_pre_match_comp, mean)
  pre_monthly_data_mean <- melt(pre_monthly_data_mean[, -c(4, 5, 6)], 
                                id.vars = c("specialty", "region", "group"))
  pre_monthly_data_mean[, "variable"] <- 
    ifelse(pre_monthly_data_mean[, "variable"] == "rel_nrx_3", -3,
           ifelse(pre_monthly_data_mean[, "variable"] == "rel_nrx_2", -2, -1))

  pre_monthly_data_mean <- dcast(pre_monthly_data_mean, 
                                 specialty + region + variable ~ group)
  colnames(pre_monthly_data_mean) <- c("specialty", "region", "period",
                                     "mean_test", "mean_control")
  #sd
  pre_monthly_data_sd <- 
    aggregate(cbind(rel_trx_3, rel_trx_2, rel_trx_1,
                    rel_nrx_3, rel_nrx_2, rel_nrx_1) ~ specialty + region +
                group, data = for_pre_match_comp, sd)

  pre_monthly_data_sd <- melt(pre_monthly_data_sd[, -c(4, 5, 6)], 
                              id.vars = c("specialty", "region", "group"))
  pre_monthly_data_sd[, "variable"] <- 
    ifelse(pre_monthly_data_sd[, "variable"] == "rel_nrx_3", -3,
           ifelse(pre_monthly_data_sd[, "variable"] == "rel_nrx_2", -2, -1))
  
  pre_monthly_data_sd <- dcast(pre_monthly_data_sd, 
                                 specialty + region + variable ~ group)
  colnames(pre_monthly_data_sd) <- c("specialty", "region", "period",
                                    "sd_test", "sd_control")
  #n
  pre_monthly_data_n <- 
    aggregate(cbind(rel_trx_3, rel_trx_2, rel_trx_1,
                    rel_nrx_3, rel_nrx_2, rel_nrx_1) ~ specialty + region +
                group, data = for_pre_match_comp, length)

  pre_monthly_data_n <- melt(pre_monthly_data_n[, -c(4, 5, 6)], 
                              id.vars = c("specialty", "region", "group"))
  pre_monthly_data_n[, "variable"] <- 
    ifelse(pre_monthly_data_n[, "variable"] == "rel_nrx_3", -3,
           ifelse(pre_monthly_data_n[, "variable"] == "rel_nrx_2", -2, -1))
  
  pre_monthly_data_n <- dcast(pre_monthly_data_n, 
                               specialty + region + variable ~ group)
  colnames(pre_monthly_data_n) <- c("specialty", "region", "period",
                                    "n_test", "n_control")
  
  pre_match_mean <- cbind(pre_monthly_data_n, pre_monthly_data_mean[, c(4, 5)],
                          pre_monthly_data_sd[, c(4, 5)])
  se <- sqrt(pre_match_mean[, "sd_test"] ^2 / pre_match_mean[, "n_test"] +
             pre_match_mean[, "sd_control"] ^2 / pre_match_mean[, "n_control"])
  pre_match_mean[, "p"] <- 
    ifelse(se > 0, 
           2 * (1 - pnorm(abs(pre_match_mean[, "mean_test"] -
                                pre_match_mean[, "mean_control"]) / se)),
            NA) 

  ## for all region
  #mean
  pre_monthly_data_mean_allr <- 
    aggregate(cbind(rel_trx_3, rel_trx_2, rel_trx_1,
                    rel_nrx_3, rel_nrx_2, rel_nrx_1) ~ specialty + group,
              data = for_pre_match_comp, mean)
  pre_monthly_data_mean_allr <- melt(pre_monthly_data_mean_allr[, -c(3,4,5)], 
                                id.vars = c("specialty", "group"))
  pre_monthly_data_mean_allr[, "variable"] <- 
    ifelse(pre_monthly_data_mean_allr[, "variable"] == "rel_nrx_3", -3,
           ifelse(pre_monthly_data_mean_allr[, "variable"] == "rel_nrx_2", -2, -1))
  
  pre_monthly_data_mean_allr <- dcast(pre_monthly_data_mean_allr, 
                                 specialty + variable ~ group)
  colnames(pre_monthly_data_mean_allr) <- c("specialty", "period",
                                       "mean_test", "mean_control")
  #sd
  pre_monthly_data_sd_allr <- 
    aggregate(cbind(rel_trx_3, rel_trx_2, rel_trx_1,
                    rel_nrx_3, rel_nrx_2, rel_nrx_1) ~ specialty+
                group, data = for_pre_match_comp, sd)
  
  pre_monthly_data_sd_allr <- melt(pre_monthly_data_sd_allr[, -c(3,4,5)], 
                              id.vars = c("specialty", "group"))
  pre_monthly_data_sd_allr[, "variable"] <- 
    ifelse(pre_monthly_data_sd_allr[, "variable"] == "rel_nrx_3", -3,
           ifelse(pre_monthly_data_sd_allr[, "variable"] == "rel_nrx_2", -2, -1))
  
  pre_monthly_data_sd_allr <- dcast(pre_monthly_data_sd_allr, 
                               specialty + variable ~ group)
  colnames(pre_monthly_data_sd_allr) <- c("specialty","period",
                                     "sd_test", "sd_control")
  #n
  pre_monthly_data_n_allr <- 
    aggregate(cbind(rel_trx_3, rel_trx_2, rel_trx_1,
                    rel_nrx_3, rel_nrx_2, rel_nrx_1) ~ specialty +
                group, data = for_pre_match_comp, length)
  
  pre_monthly_data_n_allr <- melt(pre_monthly_data_n_allr[, -c(3,4,5)], 
                             id.vars = c("specialty", "group"))
  pre_monthly_data_n_allr[, "variable"] <- 
    ifelse(pre_monthly_data_n_allr[, "variable"] == "rel_nrx_3", -3,
           ifelse(pre_monthly_data_n_allr[, "variable"] == "rel_nrx_2", -2, -1))
  
  pre_monthly_data_n_allr <- dcast(pre_monthly_data_n_allr, 
                              specialty + variable ~ group)
  colnames(pre_monthly_data_n_allr) <- c("specialty", "period",
                                    "n_test", "n_control")
  
  pre_match_mean_allr <- cbind(pre_monthly_data_n_allr, 
                               pre_monthly_data_mean_allr[, c(3,4)],
                               pre_monthly_data_sd_allr[, c(3,4)])
  se_allr <- sqrt(pre_match_mean_allr[, "sd_test"] ^2 / pre_match_mean_allr[, "n_test"] +
                  pre_match_mean_allr[, "sd_control"] ^2 / pre_match_mean_allr[, "n_control"])
  pre_match_mean_allr[, "p"] <- 
    ifelse(se_allr > 0, 
           2 * (1 - pnorm(abs(pre_match_mean_allr[, "mean_test"] -
                                pre_match_mean_allr[, "mean_control"]) / se_allr)),
           NA) 
  region <- rep("ALL", nrow(pre_match_mean_allr))
  pre_match_mean_allr <- cbind(specialty = pre_match_mean_allr[, 1],
                               region = region,
                               pre_match_mean_allr[, 2:ncol(pre_match_mean_allr)])
  pre_match_mean_total <- rbind(pre_match_mean_allr, pre_match_mean)

  # for penetration
  for_penetration <- rbind(all_test_hcps_1, control_hcps_1)
  cut_week <- total_wks + as.numeric(format(cstart, "%U")) -
    as.numeric(format(imsdate2, "%U"))+ 
    (as.numeric(format(cstart, "%Y")) - as.numeric(format(imsdate2, "%Y"))) * 52 - 1

  pre_vars <- seq(cut_week - 26, cut_week - 1)
  post_vars <- seq(cut_week, total_wks)

  var_rel_trx_pre <- paste("trx_wk_", pre_vars, sep = "")
  var_rel_trx_post <- paste("trx_wk_", post_vars, sep = "")
  
  var_rel_trx_oth_pre <- paste("trx_oth_", pre_vars, sep = "")
  var_rel_trx_oth_post <- paste("trx_oth_", post_vars, sep = "")
  
  for_penetration[, "pre_trx"] <- apply(for_penetration[, var_rel_trx_pre], 1,
                                        sum)
  for_penetration[, "post_trx"] <- apply(for_penetration[, var_rel_trx_post], 1,
                                        sum)
  for_penetration[, "pre_oth_trx"] <- apply(for_penetration[, var_rel_trx_oth_pre], 1,
                                            sum)
  for_penetration[, "post_oth_trx"] <- apply(for_penetration[, var_rel_trx_oth_post], 1,
                                             sum)
  for_penetration[, "pre_pen"] <- as.numeric(for_penetration[, "pre_trx"] > 0)
  for_penetration[, "post_pen"] <- as.numeric(for_penetration[, "post_trx"] > 0)

  for_penetration[, "pre_trx_mkt"] <-for_penetration[, "pre_trx"] +
    for_penetration[, "pre_oth_trx"]
  for_penetration[, "post_trx_mkt"] <-for_penetration[, "post_trx"] +
    for_penetration[, "post_oth_trx"]

  penetration <- aggregate(cbind(pre_pen, post_pen, pre_trx, post_trx,
                                 pre_trx_mkt, post_trx_mkt, 
                                 rep(1, nrow(for_penetration))) ~ group,
                           for_penetration,
                           sum, na.rm = TRUE)
  penetration[, "pre_penetration"] <- penetration[, "pre_pen"] / penetration[, "V7"]
  penetration[, "post_penetration"] <- penetration[, "post_pen"] / penetration[, "V7"]
  penetration[, "pre_share"] <- penetration[, "pre_trx"] / penetration[, "pre_trx_mkt"]
  penetration[, "post_share"] <- penetration[, "post_trx"] / penetration[, "post_trx_mkt"]
  
  penetration <- penetration[, c("group", "pre_penetration", "post_penetration",
                                 "pre_share", "post_share")]

  # match function
  hcp_matched_all <- NULL
  match_by_wm <- function(by = "week", match_rx){
    #by = "week"; match_rx = "nrx"
    if (by == "week") {
      .to <- n_wks
      .list <- all_wks
    } else {
      .to <- n_mth
      .list <- all_mth
    }
    
    for (eng_wks in 1 : .to) {
      from <- .list[eng_wks] - 25
      to1 <- from + 12
      from2 <- from + 13
      to2 <- from + 21
      from3 <- from + 22
      to3 <- from + 25
      
      if (by == "week") {
        hcp_week <- all_test_hcps[all_test_hcps[, "from_week"] == from, ]
      } else {
        hcp_week <- all_test_hcps[all_test_hcps[, "month"] == eng_wks, ]
      }
      
      vars_diff1 <- setdiff(colnames(hcp_week), colnames(control_hcps))
      hcp_week_1 <- hcp_week[, setdiff(colnames(hcp_week), vars_diff1)]
      hcp_week_1 <- cbind(hcp_week_1, cohort = 1, novo_test = "YES")
      control_hcps_1 <- cbind(control_hcps, cohort = 0, novo_test = "NO")
      
      hcp_rx_match <- rbind(control_hcps_1, hcp_week_1)
      
      hcp_rx_match[ , "pre_3_nov"] <- 
        apply(hcp_rx_match[, paste(match_rx, "_wk_", seq(from, to1), sep = "")],
              1, sum)
      hcp_rx_match[ , "pre_2_nov"] <- 
        apply(hcp_rx_match[, paste(match_rx, "_wk_", seq(from2, to2), sep = "")],
              1, sum)
      hcp_rx_match[ , "pre_1_nov"] <- 
        apply(hcp_rx_match[, paste(match_rx, "_wk_", seq(from3, to3), sep = "")],
              1, sum)
      
      hcp_rx_match[ , "pre_3_oth"] <- 
        apply(hcp_rx_match[, paste(match_rx, "_oth_", seq(from, to1), sep = "")],
              1, sum)
      hcp_rx_match[ , "pre_2_oth"] <- 
        apply(hcp_rx_match[, paste(match_rx, "_oth_", seq(from2, to2), sep = "")],
              1, sum)
      hcp_rx_match[ , "pre_1_oth"] <- 
        apply(hcp_rx_match[, paste(match_rx, "_oth_", seq(from3, to3), sep = "")],
              1, sum)
      
      hcp_rx_match[ , "pre_3_mktrx"] <- 
        apply(hcp_rx_match[, paste("trx", "_wk_", seq(from, to1), sep = "")],
              1, sum) +
        apply(hcp_rx_match[, paste("trx", "_oth_", seq(from, to1), sep = "")],
              1, sum)
      
      hcp_rx_match[ , "pre_2_mktrx"] <- 
        apply(hcp_rx_match[, paste("trx", "_wk_", seq(from2, to2), sep = "")],
              1, sum) +
        apply(hcp_rx_match[, paste("trx", "_oth_", seq(from2, to2), sep = "")],
              1, sum)
      
      hcp_rx_match[ , "pre_1_mktrx"] <- 
        apply(hcp_rx_match[, paste("trx", "_wk_", seq(from3, to3), sep = "")],
              1, sum) +
        apply(hcp_rx_match[, paste("trx", "_oth_", seq(from3, to3), sep = "")],
              1, sum)
      
      hcp_rx_match[ , "pre_mktrx"] <-  hcp_rx_match[ , "pre_3_mktrx"] +
        hcp_rx_match[ , "pre_2_mktrx"] + hcp_rx_match[ , "pre_1_mktrx"]
      
      hcp_rx_match[ , "pre_oth12"] <-
        hcp_rx_match[ , "pre_1_oth"] + hcp_rx_match[ , "pre_2_oth"]
      hcp_rx_match[ , "pre_nov12"] <-
        hcp_rx_match[ , "pre_1_nov"] + hcp_rx_match[ , "pre_2_nov"]
      hcp_rx_match[ , "pre_nov"] <- hcp_rx_match[ , "pre_1_nov"] + 
        hcp_rx_match[ , "pre_2_nov"] + hcp_rx_match[ , "pre_3_nov"]
      
      nov12 <- 0.1 * mean(hcp_rx_match[ , "pre_nov12"], na.rm = TRUE)
      novall <-  0.1 * mean(hcp_rx_match[ , "pre_nov"], na.rm = TRUE)
      
      if (match_rx == "nrx") {
        formula <- cohort ~ spec_grp + x_coord_albers_nbr + y_coord_albers_nbr +  
          pre_3_nov + pre_2_nov + pre_1_nov  + pre_3_oth + pre_oth12 + pre_mktrx     
      } else {
        formula <- cohort ~ spec_grp + x_coord_albers_nbr + y_coord_albers_nbr +  
          pre_3_nov + pre_2_nov + pre_1_nov  + pre_3_oth + pre_oth12 
      }
      hcp_rx_match1 <- hcp_rx_match
      hcp_rx_match1[, "spec_grp"] <- as.factor(hcp_rx_match1[, "spec_grp"])
      ps_model <- glm(formula, data = hcp_rx_match1, family = "binomial")
      rm(hcp_rx_match1)
      gc()
      
      hcps_rx_4match <- cbind(hcp_rx_match, pscore = fitted(ps_model))
      
      # for match
      .trgind <- c("spec_grp", "reg_num", "lat_nbr", "lon_nbr",
                   "pre_nov12", "pre_nov", "pscore")
      .wts <- c(1, 1, 1, 1, 5, 5, 5)
      .mk <- c(0, 0, 1, 1, nov12, novall, 0.001)
      hcp_matched <- ps_match(data    = hcps_rx_4match, 
                              group   = "cohort", 
                              id	    = "imsdr7", 
                              mvars   = .trgind, 
                              wts	    = .wts ,
                              dmaxk   = .mk, 
                              transf  = 0,
                              dist    = 1, 
                              ncontls = 1,
                              seedca  = 279,
                              seedco  = 389)
     assign(paste("matched_week_", from, sep = ""),
            cbind(hcp_matched[[1]], engage_wk = from + 25)) 
     matched_c <- subset(get(paste("matched_week_", from, sep = "")), cohort == 0)
     control_hcps <- subset(control_hcps, !imsdr7 %in% unique(matched_c[, "imsdr7"]))
     hcp_matched_all <- rbind(hcp_matched_all, 
                              get(paste("matched_week_", from, sep = "")))                     
    }
    hcp_matched_all <- 
      left_join(hcp_matched_all,
                campaign_hcps[, c("imsdr7", "engagement_date")],
                by = "imsdr7")
    return(hcp_matched_all)
  }
  hcp_matched_all <- match_by_wm(by = "week", match_rx ="nrx")
 
  colnames(hcp_matched_all) <- tolower(colnames(hcp_matched_all))
  line4 <- length(unique(hcp_matched_all[hcp_matched_all[, "cohort"] == 1, "imsdr7"]))
  line5 <- length(unique(hcp_matched_all[hcp_matched_all[, "cohort"] == 0, "imsdr7"]))
  
  column1 <- c("Total # of Exposed HCPs", 
               paste("Campaign Exposed HCPs with ", campaign_prd, " Rx", sep = ""),
               "Total non-exposed HCPs (Potential Controls)",
               paste("Total ", campaign_prd," Writers", sep = ""),
               "Matched Pairs (1:1 matched)")
  hcp_count <- c(line1, line1x, line2x, writers, line5)
  attrition <- data.frame(metric = column1, hcp_count = hcp_count)
  

  return(list(matched_result = hcp_matched_all,
              descriptive = descriptives,
              pre_period_summary = pre_match_mean_total,
              penetration = penetration,
              attrition = attrition))
}


# call the main function for part1
system.time(part1 <- part1_pre_process(raw_xpo_data = "ims_rx_data_sample1",
                                       campaign_prd = "VICTOZA",
                                       pre_wks = 26,
                                       campaign_data = "campaign_139",
                                       study_rx = "nrx",
                                       total_wks = 128))

# output the matched results as the input of part2
write.csv(part1[[1]], "hcp_matched_all.csv", row.names = FALSE)
  


