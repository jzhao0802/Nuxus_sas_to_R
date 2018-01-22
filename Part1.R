#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#   Project: NEXXUS
#   Purpose: Loading and Cleaning data for Propensity Score Matching
#   Version: 0.1
#   Programmer: Xin Huang
#   Date: 09/09/2015
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# set up the referrence and load some required library
setwd("D:\\Projects\\Nexxus")

install.packages("Reshape2")
require(ggplot2)
require(dplyr) # merge
require(reshape2) #transponse
require(sas7bdat)
require(stringr) # pad zero in ids

prog = NULL
min_postwks = 4


# %let  Raw_Xpo_data = ims_rx_data_Sample;          
# %let	Campaign_PRD = VICTOZA;     
# %let	Pre_wks      = 26;          
# %let	Campaign_dat = Campaign_139; 
# %let	Study_RX     = NRX;
# %let	Total_wks 	 = 128;

# read in the sas data
ddHeader <- tolower(scan("chk.csv", what=character(), sep=',', nlines=1))
cc <- c(imsdr7 = "character",
        zipcode = "character",
        spec = "character",
        specialty = "character",
        product = "character",
        datadate = "character",
        week = "numeric",
        trx = "numeric",
        nrx = "numeric")

ims_rx_data_sample <- read.csv("chk.csv", col.names = ddHeader, colClasses=cc,
                               stringsAsFactors = FALSE)

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
  
  # returns:
  #  matched data
  
#   raw_xpo_data = "ims_rx_data_sample"
#   campaign_prd = "VICTOZA"
#   pre_wks = 26
#   campaign_data = "campaign_139"
#   study_rx = "nrx"
#   total_wks = 128
  infile = "raw_xpo_data"
  min_postwks <- 4
  
  raw_data <- subset(get(get(infile)), as.numeric(imsdr7) < 9000000)
  raw_data[, "brand"] <- toupper(raw_data[, "product"])
  level <- c("imsdr7", "spec", "zipcode", "brand")
  aggregated <-
    aggregate(cbind(nrx, trx) ~ imsdr7 + spec + zipcode + brand + week, 
              raw_data, sum, na.rm = TRUE)
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
  
  tmp[, "brand"] <-ifelse(tmp[, "brand"] == campaign_prd, tmp[, "brand"], "OTHER")
  tmp[, c("lat_nbr", "lon_nbr", "x_coord_albers_nbr","y_coord_albers_nbr")] <-
    lapply(tmp[, c("lat_nbr", "lon_nbr", "x_coord_albers_nbr","y_coord_albers_nbr")],
           function(x) ifelse(is.na(x), 0, x))
  
  # do the summary at level imsdr7 + spec + zipcode + brand
  
  assign(gsub(" ", "",paste("ims_xpo_dat_", tolower(campaign_prd), "_rx")), 
         aggregate(.~imsdr7 + spec + zipcode + brand + spec_grp + region + 
                     lat_nbr + lon_nbr + x_coord_albers_nbr + y_coord_albers_nbr,
                   tmp[, setdiff(colnames(tmp), c("grp_id", "state"))],
                   sum,
                   na.action = na.pass))
  
  # read in the campaigm data
  ddHeader2 <- tolower(scan(campaign_data, what=character(), 
                            sep=',', nlines=1))
  cc2 <- c(contactid = "character",
           engagement_date = "character",
           camapign_date = "character",
           campaignid = "character",
           imsid = "character")
  campaign_hcps <- read.csv(campaign_data, col.names = ddHeader2,
                            colClasses = cc2,
                            stringsAsFactors = FALSE)
  
  campaign_hcps[, "imsdr7"] <- str_pad(campaign_hcps[, "imsid"], 7, pad = "0")
  campaign_hcps <- subset(campaign_hcps, imsdr7 != "0000000")
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
  
  ims_rx_1rec_hcp <- full_join(ims_rx_1rec_nov, ims_rx_1rec_oth, 
                               by =c("imsdr7", "spec", "zipcode",
                                     "brand", "spec_grp", "region", "lat_nbr",
                                     "lon_nbr", "x_coord_albers_nbr",
                                     "y_coord_albers_nbr"))
  
  replace_na <- function(x) {
    x[is.na(x)] <- 0
    x
  }
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
  outlier <- function(){
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
  hcps_rx_all <- outlier()
  
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
  
  all_wks <- paste(weeks[, "week"], collapse = "*")
  n_wks <- length(weeks[, "week"])
  
  all_mth <- paste(weeks[, "month"], collapse = "*")
  n_mth <- length(weeks[, "month"])
  
  # for test dataset
  all_test_hcps <- inner_join(hcps_rx_all, test_hcp, by = "imsdr7")
  all_test_hcps <- subset(all_test_hcps, eligible == 1)
  
  all_test_hcps[, "sr_elig"] <- ifelse(is.na(all_test_hcps[, "spec"]) & 
                                         is.na(all_test_hcps[, "region"]), 
                                       0, 1)
  all_test_hcps[, "spec_num"] <- all_test_hcps[, "spec_grp"]
  all_test_hcps[, "region"] <- ifelse(is.na(all_test_hcps[, "region"]), 
                                      "UNKNOWN", all_test_hcps[, "region"])
  all_test_hcps[, "reg_num"] <- 
    (all_test_hcps[, "region"] == "MIDWEST") * 1 +
    (all_test_hcps[, "region"] == "NORTHEAST") * 2 +
    (all_test_hcps[, "region"] == "SOUTH") * 3 +
    (all_test_hcps[, "region"] == "WEST") * 4
  
  # for control group
  control_hcps <- anti_join(hcps_rx_all, test_hcp, by = "imsdr7")
  
  control_hcps[, "sr_elig"] <- ifelse(is.na(control_hcps[, "spec"]) & 
                                         is.na(control_hcps[, "region"]), 
                                       0, 1)
  control_hcps[, "spec_num"] <- control_hcps[, "spec_grp"]
  control_hcps[, "region"] <- ifelse(is.na(control_hcps[, "region"]), 
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
  
  
  descriptive_dat[, "post_nov"] <- 
    descriptive_dat[, "post_nov"] * pre_wks / post_len
  descriptive_dat[, "post_nov_nrx"] <- 
    descriptive_dat[, "post_nov_nrx"] * pre_wks / post_len
  descriptive_dat[, "post_oth"] <- 
    descriptive_dat[, "post_oth"] * pre_wks / post_len
  descriptive_dat[, "post_oth_nrx"] <-
    descriptive_dat[, "post_oth_nrx"] * pre_wks / post_len
  
  descriptive_dat[, "pre_writer_x"] <-  descriptive_dat[, "pre_nall"] > 0
  descriptive_dat[, "pre_writer_m"] <- 
    descriptive_dat[, "pre_nall"] > 0 | descriptive_dat[, "pre_oall"] > 0
  descriptive_dat[, "pre_writer_c"] <-  descriptive_dat[, "pre_oall"] > 0
  
  descriptive_dat[, "post_writer_x"] <-  descriptive_dat[, "post_nall"] > 0
  descriptive_dat[, "post_writer_m"] <- 
    descriptive_dat[, "post_nall"] > 0 | descriptive_dat[, "post_oall"] > 0
  descriptive_dat[, "post_writer_c"] <-  descriptive_dat[, "post_oall"] > 0
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
    
    
  
  
  
  
  
  
  
  
  
  
  
                         
  
  
  
  




}


