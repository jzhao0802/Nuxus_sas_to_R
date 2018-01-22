#*********************************************************************************************
#NEXXUS Process Part-2
#This is the ROI analysis part. It starts from the matched pair created in Part-1 process.
#This code will run ANCOVA for product x, competitor products and market Rx and 
#prepare output for all designed metrics. THe out put tables will be written in excel file 
#with tabs:
#    Attrition, Descriptive_Table, Market Penetration, ANCOVA_Test, PreMatch_Test_Ctrl
#and PostMatch_Test_Ctrl, Monthly_Rx (For pre-post Trend)


#Final excel file will be saved in Out_dir with file name: OutPut_Campaign_&Prog_num..xlsx

#Developer : Jie Zhao
#*********************************************************************************************;




# requeire libraries
library(sas7bdat)
library(zoo)
library(xlsx)
library(plyr)
library(dplyr)
library(lsmeans)
library(reshape)
library(snow)
library(snowfall)

# needed file path defination
dataPath <- paste("\\\\plyvnas01\\statservices2\\CustomStudies\\Promotion Management", 
                  "\\2014\\NOVOLOG\\04 Codes\\Nexxus\\R_Version\\1016", sep='')
outPath <- paste("\\\\plyvnas01\\statservices2\\CustomStudies\\Promotion Management", 
                 "\\2014\\NOVOLOG\\04 Codes\\Nexxus\\R_Version\\1016", sep='')

Raw_Xpo_data = 'for_sas_test_0929' # Xponent Raw Rx Data
    
Ext <- '.csv'
Ext1 <- '.sas7bdat'

part2_function <- function(Campaign_PRD, Pre_wks, Target_RX, 
                           Total_wks, Campaign_dat, Specfile,
                           ZIP2Region, n.cpu){
  
  #@@Parameters:
    #Campaign_PRD = 'VICTOZA' # Campaign Product Name
    #Pre_wks      = 26        #Same as defined in part-1 
    #Target_RX	  = 'NRX'     #Match and analysis Rx. Same as used in Part-1
    #Total_wks    = 128       #Total available data weeks. Same as used in Part-1 
    #Campaign_dat = 139       #Campaign Contact HCP list excel file 
    #Specfile	    = ''        #Excel file for customized specialty groups. Default = All in 1
    #ZIP2Region	  =  ''       #Excel file for Zip to customized region. Defaulat=Census regions);
  #@@Returns: 
    #A list contains what the SAS outputs
    
#   Campaign_PRD = 'VICTOZA'
#   Pre_wks=26
#   Target_RX='NRX'
#   Total_wks=128
#   Campaign_dat=139
#   Specfile=''
#   ZIP2Region=''
  
  #Campaign Contact HCP list excel file 
  Campaign_dat <- paste('CAMPAIGN', Campaign_dat, sep='_')
  
  #read in input data
  campId <- 
    as.vector(read.csv(paste(dataPath, '\\',
                             Campaign_dat, '_id',
                             Ext, sep=''))[1,1])
 
  CStart <-
    min(as.Date(as.character(read.csv(paste(dataPath, '\\Campaign_HCP_',
                                            campId, Ext, sep=''))$Engagement_Date),
                format='%d%b%Y'))  
  
  names <- scan(file=paste(dataPath, '\\', Raw_Xpo_data, Ext, sep=''),
                nlines=1, what="complex", sep=',', skip=0)
  line1 <- scan(file=paste(dataPath, '\\', Raw_Xpo_data, Ext, sep=''),
                nlines=1, what="complex", sep=',', skip=1)
  date_num <- line1[6]
  
  from <- Total_wks + as.numeric(strftime(CStart,format="%W")) +
    (as.numeric(strftime(CStart,format="%Y")) - 
       as.numeric(strftime(line1[names=='datadate'],format="%Y"))) *52 - 
    as.numeric(strftime(line1[names=='datadate'],format="%W")) - 26
  to <- Total_wks + as.numeric(strftime(CStart,format="%W")) +
    (as.numeric(strftime(CStart,format="%Y")) - 
       as.numeric(strftime(line1[names=='datadate'],format="%Y"))) *52 - 
    as.numeric(strftime(line1[names=='datadate'],format="%W")) - 1
  IMSDate2 <- line1[names=='datadate']
  
  #read in the sasdatax.HCP_MATCHED_ALL_&Prog_num.
  matched_data <- read.table(paste(dataPath, '\\hcp_matched_all', Ext, sep=''),
                             sep=',', header=T, stringsAsFactors=FALSE)
  # dim(matched_data) #[1] 15940   534
  colnames(matched_data)<- tolower(colnames(matched_data))
  var_list <- colnames(matched_data)
  
  # check the covariates existence
  covar_nece <- c("imsdr7",          
                  "region",          
                  "engagement_date", 
                  "cohort",          
                  "engage_wk")
  type <- c("integer",        
            "character",        
            "character",       
            "integer",       
            "integer")
  
  if (any(is.na(match(covar_nece, var_list)))){
      covar_nece_str <- paste(covar_nece, sep='', collapse=" ")
      stop(paste("Please make sure that the variables below are all included in your matched data!\n", covar_nece_str, '\n', sep=''))
  }
  if (any(sapply(matched_data[1:100, covar_nece], class)!=type)){
      type_str <- paste(rbind(paste(covar_nece, ':', sep=''), type), sep='', collapse=' ')
      stop(paste("the covariates data type is not matched the following expectation!\nt", type_str, '\n', sep=''))
  }
  #covar checking end
  
  
  matched_data1 <- matched_data[, match(grep('^(imsdr7|engage|trx|nrx)|cohort$',
                                             var_list, value=T), var_list)]
  var_list1 <- colnames(matched_data1)
  #rm(list=c('matched_data'))
  
  othT <- grep('^trx_oth', var_list1, value=T)
  othN <- grep('^nrx_oth', var_list1, value=T)
  nlT <- grep('^trx_wk', var_list1, value=T)
  nlN <- grep('^nrx_wk', var_list1, value=T)
  #matched_data1_test <- matched_data1[1:100,]
  #var_list1[apply(matched_data1_test, 2, is.numeric)]
  
  matched_data2_temp <- function(r){
    x <- matched_data1[r, ]
    if (x[match('cohort', var_list1)]==1){
        group <- 'TEST'
    }else{
        group <- 'CONTROL'
    }
    
    temp1 <- lapply(1:Total_wks, function(i){
        #gap <- i - x[match('engage_wk', var_list1)] - 1
        gap <- i - x$engage_wk - 1
        if(gap > (-1-Pre_wks) & gap < 0){
            pre_flag <- 1
        }else if(gap > 0){
            pre_flag <- 0
        }else{
            pre_flag <- -1
        }
        return(list(pre_flag, gap))
    })
    library(plyr)
    temp1_1 <- ldply(temp1, quickdf)
    pre_flag <- temp1_1[, 1]
    gap <- temp1_1[, 2][Total_wks]
    pre_othT_withMiss <- x[match(othT[pre_flag==1], var_list1)]
    pre_nlT_withMiss <- x[match(nlT[pre_flag==1], var_list1)]
    pre_othN_withMiss <- x[match(othN[pre_flag==1], var_list1)]
    pre_nlN_withMiss <- x[match(nlN[pre_flag==1], var_list1)]
    
    pre_period_othT <- sum(replace(pre_othT_withMiss, 
                                   which(is.na(pre_othT_withMiss)), 0))
    pre_period_nlT <- sum(replace(pre_nlT_withMiss, 
                                  which(is.na(pre_nlT_withMiss)), 0))
    pre_period_othN <- sum(replace(pre_othN_withMiss,
                                   which(is.na(pre_othN_withMiss)), 0))
    pre_period_nlN <- sum(replace(pre_nlN_withMiss, 
                                  which(is.na(pre_nlN_withMiss)), 0))
    pre_market_T <- pre_period_othT + pre_period_nlT
    pre_market_N <- pre_period_othN + pre_period_nlN
    
    post_othT_withMiss <- x[match(othT[pre_flag==0], var_list1)]
    post_nlT_withMiss <- x[match(nlT[pre_flag==0], var_list1)]
    post_othN_withMiss <- x[match(othN[pre_flag==0], var_list1)]
    post_nlN_withMiss <- x[match(nlN[pre_flag==0], var_list1)]
    
    post_period_othT <- sum(replace(post_othT_withMiss, 
                                    which(is.na(post_othT_withMiss)), 0))
    post_period_nlT <- sum(replace(post_nlT_withMiss,
                                   which(is.na(post_nlT_withMiss)), 0))
    post_period_othN <- sum(replace(post_othN_withMiss, 
                                    which(is.na(post_othN_withMiss)), 0))
    post_period_nlN <- sum(replace(post_nlN_withMiss,
                                   which(is.na(post_nlN_withMiss)), 0))
    post_market_T <- post_period_othT + post_period_nlT
    post_market_N <- post_period_othN + post_period_nlN
    
    temp2 <- c(pre_period_othT=pre_period_othT, 
               pre_period_nlT=pre_period_nlT, 
               pre_period_othN=pre_period_othN,  
               pre_period_nlN=pre_period_nlN,    
               pre_market_T=pre_market_T,  
               pre_market_N=pre_market_N,
               post_period_othT=post_period_othT,
               post_period_nlT=post_period_nlT,
               post_period_othN=post_period_othN, 
               post_period_nlN=post_period_nlN, 
               post_market_T=post_market_T,   
               post_market_N=post_market_N)
    temp2_1=c(temp2[1:6] * Pre_wks/sum(pre_flag==1, na.rm=T),
              temp2[7:12] * Pre_wks/sum(pre_flag==0, na.rm=T))
    x1 <- x[-match(c('engage_wk'), var_list1)]
    x2 <- as.vector(t(x1))
    names(x2) <- names(x1)
    return(c(x2, group=group, gap=gap, temp2_1))
}
start_time<- proc.time()
num_pros <- n.cpu
sfInit(parallel=TRUE, cpus=num_pros)
sfLibrary("dplyr",character.only = TRUE)

sfExport('matched_data1'
         ,'othT'
         ,'othN'
         ,'nlT'
         , 'nlN'
         , 'var_list1'
         , 'Total_wks'
         , 'Pre_wks')
sfExport("ldply", namespace = "plyr")

#run the function thru the combos
#matched_data2_1 <- sfClusterApplyLB(1:nrow(matched_data1), matched_data2_temp)
matched_data2_1 <- sfClusterApplyLB(1:nrow(matched_data1), matched_data2_temp)
sfStop()
end_time<- proc.time()
matched_data2 <- ldply(matched_data2_1, quickdf)
#lapply(matched_data2_1, length)

  
  
  matched_data3_1 <- 
    as.data.frame(apply(matched_data2[, -match(c('engagement_date', 'group'),
                                               names(matched_data2))], 2, as.numeric))
  matched_data3 <- 
    cbind(matched_data3_1, group=matched_data2$group, 
          engagement_date=matched_data2$engagement_date)

  #Get HCP counts for matched Test and controls after check min post weeks 
  line6 <- length(unique(matched_data3[matched_data3$cohort==1, 'imsdr7']))
  line7 <- length(unique(matched_data3[matched_data3$cohort==0, 'imsdr7']))
  
  #Get the adjusted Post RX counts using Proc GLM or MIXED - the ANCOVA analysis
  options("contrasts")
  product_list <- c('Product X            ', 
                    'Product Other than X ',
                    'Market (All Products)')
  esti_list <- lapply(c('post_period_nlN', 'post_period_othN', 'post_market_N'),
                      function(v){
                        pre_v <- gsub('post', 'pre', v)
                        eval(parse(text=paste('results = with(matched_data3, 
                                              lm(', v, ' ~ group + ', pre_v, ' ))',
                                              sep='' )))
                        #summary <- summary.lm(results)
                        esti <- summary(lsmeans(results, "group"))[, 2] 
                        matched_hcps <- line6
                        index <- esti[2]/esti[1]
                        p_value <- 
                         1- pnorm((summary(lsmeans(results, "group"))[2, 2] -
                                    summary(lsmeans(results, "group"))[1, 2]) /
                                    (sqrt((summary(lsmeans(results, "group"))[2, 3]) ^ 2 +
                                            (summary(lsmeans(results, "group"))[1, 3]) ^ 2)
                                    )) 
                        gains <- esti[2]-esti[1]
                        return(c(product=
                                   product_list[match(v, 
                                                      c('post_period_nlN',
                                                        'post_period_othN', 
                                                         'post_market_N'))],
                                 matched_hcps=matched_hcps, test=unname(esti[2]),
                                 cont=unname(esti[1]), 
                                 gains=unname(gains), 
                                 index=unname(index),
                                 p_value=unname(p_value)))
                    })
  esti_df <- ldply(esti_list, quickdf)
  
  
  # START: create points for graph
  # matched_data_forGp 
get_graph_data_1 <- function(r){
    x <- matched_data[r, ]
    if(x$cohort==1){
        group <- 'TEST'
    }else{
        group <- 'CONTROL'
    }
    
    gap_vct <- c(1:Total_wks) - x$engage_wk - 1 
    gap_flag <- ifelse(gap_vct >(-1-Pre_wks) & gap_vct < (1+Pre_wks), T, F)
    gap_flag0 <- 
        ifelse(gap_vct >(-1-Pre_wks) & gap_vct < (1+Pre_wks) & gap_vct > 0, 
               1, 0)
    breaks <- c(c(-Pre_wks+c(-Inf, -1, 4, 8, 13, 17, 21, 26)),
                c(4, 8, 13, 17, 21, 26))
    Rel_month_vct <- c(-Inf, seq(-6, -1, 1), seq(1, 6, 1))
    bucket <- as.numeric(as.vector(cut(gap_vct, breaks, right=F, 
                                       labels=Rel_month_vct)))
    rel_month <- bucket[gap_flag]
    post_n <- sum(gap_flag0)
    
    breaks1 <- c(-Inf, -Pre_wks-0.1, -14, -5, -1, Inf)
    bucket_period <-
        as.numeric(as.vector(cut(gap_vct[gap_flag], 
                                 breaks1, right=T,
                                 labels=c(-999, -3, -2, -1, 999))))
    period <- ifelse(bucket_period==999 | bucket_period==-999, 
                     rel_month, bucket_period)
    rel_trx <- as.vector(t(x[match(nlT, var_list)][gap_flag]))
    rel_nrx <- as.vector(t(x[match(nlN, var_list)][gap_flag]))
    rel_trx_other <- as.vector(t(x[match(othT, var_list)][gap_flag]))
    rel_nrx_other <- as.vector(t(x[match(othN, var_list)][gap_flag]))
    # rel_trx <- ifelse(gap_flag, rel_trx1, NA)
    #rel_nrx <- ifelse(gap_flag, rel_nrx1, NA)
    #rel_trx_other <- ifelse(gap_flag, rel_trx_other1, NA)
    #rel_nrx_other <- ifelse(gap_flag, rel_nrx_other1, NA)
    mkt_trx <- rel_trx + rel_trx_other
    mkt_nrx <- rel_nrx + rel_nrx_other
    temp <- cbind(imsdr7=x$imsdr7, specialty= "ALL SPECIALTY",
                  group=group, rel_wk=gap_vct[gap_flag], 
                  rel_month=rel_month, rel_trx=rel_trx, 
                  rel_nrx=rel_nrx, rel_trx_other=rel_trx_other, 
                  rel_nrx_other=rel_nrx_other, mkt_trx=mkt_trx, 
                  mkt_nrx=mkt_nrx, period=period, 
                  region=as.character(x$region))
    return(temp)
    
}
start_time<- proc.time()
num_pros <- n.cpu
sfInit(parallel=TRUE, cpus=num_pros)

sfExport('matched_data','othT','othN','nlT', 'nlN', 'var_list', 'Total_wks', 'Pre_wks')

#run the function thru the combos
graph_data2_temp <- sfClusterApplyLB(1:nrow(matched_data), get_graph_data_1)
#graph_data2_temp <- sfClusterApplyLB(1:2000, get_graph_data_1)
sfStop()
end_time<- proc.time()

check_dim <- ldply(lapply(graph_data2_temp, 
                          function(i){
                              return(dim(i))
                          }), 
                   quickdf)

graph_data_2 <- ldply(graph_data2_temp, rbind)
var_list2 <- names(graph_data_2)


  #create 'speciatly' and 'region' if 'Specfile' and 'ZIP2Region' files are given.
  if(Specfile != ''){
      spec_dic <- 
        read.table(paste(dataPath, '\\', Specfile, Ext, sep=''), sep=',', header=T)
      graph_data_2$speciatly <-
        ifelse(!is.na(match(matched_data$spec, spec_dic$SPEC_CD)), 
               as.character(spec_dic[match(matched_data$spec, spec_dic$SPEC_CD),
                                     'CLASS_DESC']), '')
  }else{
      graph_data_2$specialty <- "ALL SPECIALTY"
  }
  if(ZIP2Region != ''){
      reg_dic <- 
        read.table(paste(dataPath, '\\', ZIP2Region, Ext, sep=''), sep=',', header=T)
      graph_data_2$region <-
        ifelse(!is.na(match(matched_data$zipcode, reg_dic$ZIPCODE)), 
               as.character(reg_dic[match(matched_data$zipcode, reg_dic$ZIPCODE),
                                    "REGION"]), '')
  }   
  
  character_var <- c('imsdr7', 'specialty', 'group', 'region')
  graph_data_3 <- cbind(apply(graph_data_2[, -match(character_var, var_list2)],
                              2, 
                              function(x){
                                as.numeric(as.vector(x))
                                }),
                        graph_data_2[, character_var])

  
  #    ***********  Get Monthly summary ****************;
  target_trx <- grep('trx|nrx', var_list2, value=T)
  graph_data_summary <- aggregate(graph_data_3[,target_trx], 
                                  by=list(group=graph_data_3$group,
                                          rel_month=graph_data_3$rel_month),
                                  sum)
  hcp_cnt <- aggregate(graph_data_3$imsdr7,
                       by=list(group=graph_data_3$group,
                               rel_month=graph_data_3$rel_month),
                       function(x)length(unique(x)))
  
  graph_data_summary_1 <- 
    left_join(graph_data_summary[order(graph_data_summary$group), ],
              hcp_cnt, by=c("group" = 'group', "rel_month" = "rel_month"))
  names(graph_data_summary_1)[length(graph_data_summary_1)] <- 'hcp_cnt'

  #*******************  Get Post Match test/control comparison ****************;
  monthly_data <- aggregate(cbind(graph_data_3[,target_trx], 
                                  freq=rep(1, nrow(graph_data_3))),
                            by=list(specialty=graph_data_3$specialty,
                                    region=graph_data_3$region, 
                                    group=graph_data_3$group,
                                    imsdr7= graph_data_3$imsdr7, 
                                    period=graph_data_3$period), 
                            function(x)sum(x, na.rm=T))

  monthly_data <- monthly_data[monthly_data$period<0,]
  monthly_mean <- aggregate(monthly_data[, c(target_trx)],
                            by=list(specialty=monthly_data$specialty, 
                                    region=monthly_data$region, 
                                    group=monthly_data$group, 
                                    period=monthly_data$period), 
                            function(x){mean(x, na.rm=T)})
  monthly_std <- aggregate(monthly_data[, target_trx],
                           by=list(specialty=monthly_data$specialty, 
                                   region=monthly_data$region, 
                                   group=monthly_data$group, 
                                   period=monthly_data$period),
                           function(x){sd(x, na.rm=T)})
  names(monthly_std)[match(target_trx, names(monthly_std))] <- 
    paste(names(monthly_std)[match(target_trx, names(monthly_std))], '_std', sep='')
  monthly_mean_std <- 
    left_join(monthly_mean,
              monthly_std, 
              by=c('specialty', 'region','group', 'period'))
  monthly_mean_std$freq <- 
    aggregate(rep(1, nrow(monthly_data)), 
              by=list(specialty=monthly_data$specialty, 
                      region=monthly_data$region,
                      group=monthly_data$group,
                      period=monthly_data$period),
              function(x){sum(x, na.rm=T)})[,5]
  monthly_all_region_mean <- 
    aggregate(monthly_data[, target_trx],
              by=list(specialty=monthly_data$specialty, 
                      group=monthly_data$group,
                      period=monthly_data$period),
              function(x){mean(x, na.rm=T)})
  monthly_all_region_std <- 
    aggregate(monthly_data[, target_trx], 
              by=list(specialty=monthly_data$specialty,
                      group=monthly_data$group, 
                      period=monthly_data$period), 
              function(x){sd(x, na.rm=T)})
  names(monthly_all_region_std)[match(target_trx, names(monthly_all_region_std))] <-
    paste(target_trx, '_std', sep='')
  monthly_all_region_freq <-
    aggregate(rep(1, nrow(monthly_data)), 
              by=list(specialty=monthly_data$specialty,
                      group=monthly_data$group, 
                      period=monthly_data$period), 
              function(x){sum(x, na.rm=T)})
  monthly_all_region <-
    left_join(left_join(monthly_all_region_mean,
                        monthly_all_region_std, 
                        by=c('specialty', 'group', 'period')),
              monthly_all_region_freq, by=c('specialty', 'group', 'period'))
  monthly_all_region$region <- 'all region'
  names(monthly_all_region)[match('x', names(monthly_all_region))] <- 'freq'
  monthly_mean_std_all <- rbind(monthly_all_region, monthly_mean_std)
 
  reshape <- function(input){
      levels <- levels(input$group)
      stack <- function(i){
          data <- input[input$group==i, ]
          transp_list <- lapply(1:nrow(data), function(r){
              x <- data[r,]
              temp1 <- x[c(1, 3, ncol(monthly_mean_std_all))]
              temp2 <- t(x[-c(1, 2, 3, ncol(monthly_mean_std_all))])
              new_names <- rownames(temp2)
              temp3 <- cbind(temp1, new_names,temp2)
              colnames(temp3)[ncol(temp3)] <-  i
              return(temp3)
          })
          stacked_data <- ldply(transp_list, rbind)
          return(stacked_data)
      }
      cont <- stack(levels[levels=='CONTROL'])
      test <- stack(levels[levels=='TEST'])
      cont_test <- left_join(cont, 
                             test, 
                             by=c('specialty', 'region', 'period', 'new_names'))
      return(cont_test)
  }
  pre_compare <- reshape(monthly_mean_std_all)
  var_list3 <- names(pre_compare)
  for_compare <- function(input, var){
      temp <- pre_compare[pre_compare$new_names==var, ]
      temp_ord <- temp[order(temp[, match('specialty',var_list3)], temp[, match('region', var_list3)], temp[, match('period', var_list3)]), match(c('specialty', 'region', "period", 'CONTROL', 'TEST'), var_list3)]
  }
  nov_rx <- for_compare(pre_compare, 'rel_nrx')
  nov_std <- for_compare(pre_compare, 'rel_nrx_std')
  freq <- for_compare(pre_compare, 'freq')
  levels <- grep('^CONTROL|^TEST', names(nov_std), ignore.case=T, value=T)
  names(nov_std)[match(levels, names(nov_std))] <- 
    paste(names(nov_std)[match(levels, names(nov_std))], '_std', sep='')
  names(freq)[match(levels, names(freq))] <-
    paste(names(freq)[match(levels, names(freq))], '_cnt', sep='')
  post_match_compare <- left_join(left_join(nov_rx,
                                            nov_std,
                                            by=c('specialty', 'region', 'period')),
                                  freq,
                                  by=c('specialty', 'region', 'period'))
  
  post_match_compare$p_value <- 
    unlist(lapply(1:nrow(post_match_compare),
                  function(r){
                    x <- post_match_compare[r, ]
                    for_p_value <- ifelse(x$TEST_std > 0, abs(x$TEST-x$CONTROL)/sqrt(x$TEST_std^2/x$TEST_cnt+x$CONTROL_std^2/x$CONTROL_cnt), NA)
                    p_value <- ifelse(!is.na(for_p_value), 2*(1-pnorm(for_p_value)), NA)
                    return(p_value)
                  }))
  return(list(esti_df = esti_df,
              pre_compare = pre_compare, 
              post_match_compare = post_match_compare, 
              graph_data_summary_1 = graph_data_summary_1))
}

system.time(part2_result <- part2_function(Campaign_PRD = 'VICTOZA', 
                               Pre_wks=26, 
                               Target_RX='NRX', 
                               Total_wks=128, 
                               Campaign_dat=139, 
                               Specfile='', 
                               ZIP2Region='', n.cpu=4))


esti_df <- part2_result$esti_df
pre_compare <- part2_result$pre_compare
post_match_compare <- part2_result$post_match_compare
graph_data_summary_1 <- part2_result$graph_data_summary_1



#write the result of part1
write.xlsx(part1$descriptive,
           paste(outPath, '\\OutPut_Campaign_', campId, '.xlsx', sep=''), 
           sheetName='descriptive', 
           append=T, 
           row.names=F)
write.xlsx(part1$pre_period_summary, 
           paste(outPath, '\\OutPut_Campaign_', campId, '.xlsx', sep=''),
           sheetName='pre_match_compare',
           append=T,
           row.names=F)
write.xlsx(part1$penetration, 
           paste(outPath, '\\OutPut_Campaign_', campId, '.xlsx', sep=''), 
           sheetName='penetration', 
           append=T, 
           row.names=F)
write.xlsx(part1$attrition,
           paste(outPath, '\\OutPut_Campaign_', campId, '.xlsx', sep=''),
           sheetName='attrition', 
           append=T, 
           row.names=F)

#write the result of part2
write.xlsx(esti_df, 
           paste(outPath, '\\OutPut_Campaign_', campId, '.xlsx', sep=''), 
           sheetName='anova_test',
           append=T, 
           row.names=F)
write.xlsx(pre_compare,
           paste(outPath, '\\OutPut_Campaign_', campId, '.xlsx', sep=''),
           sheetName='pre_compare', 
           append=T, 
           row.names=F)
write.xlsx(post_match_compare,
           paste(outPath, '\\OutPut_Campaign_', campId, '.xlsx', sep=''),
           sheetName='post_match_compare',
           append=T, 
           row.names=F)
write.xlsx(graph_data_summary_1,
           paste(outPath, '\\OutPut_Campaign_', campId, '.xlsx', sep=''), 
           sheetName='graph_data', 
           append=T, 
           row.names=F)

