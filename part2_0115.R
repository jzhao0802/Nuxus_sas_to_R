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
                "\\2014\\NOVOLOG\\04 Codes\\Nexxus\\R_Version\\1016\\withoutSpec", sep='')

# dataPath <- paste('D:\\jzhao\\Nexxus')
# outPath <- paste('D:\\jzhao\\Nexxus\\output')
Raw_Xpo_data = 'for_sas_test_0929' # Xponent Raw Rx Data
    
Ext <- '.csv'
Ext1 <- '.sas7bdat'

part2_function <- function(Campaign_PRD, Pre_wks, Target_RX, 
                           Total_wks, CampaignId, Specfile,
                           ZIP2Region, n.cpu){
    
    #@@Parameters:
    Campaign_PRD = 'VICTOZA' # Campaign Product Name
    Pre_wks      = 26        #Same as defined in part-1 
    Target_RX	  = 'NRX'     #Match and analysis Rx. Same as used in Part-1
    Total_wks    = 128       #Total available data weeks. Same as used in Part-1 
    #Campaign_dat = 139       #Campaign Contact HCP list excel file 
    campId <- '139'
    Specfile	    = 'SPEC_Grp'        #Excel file for customized specialty groups. Default = All in 1
    ZIP2Region	  =  'ZIP_region'       #Excel file for Zip to customized region. Defaulat=Census regions);
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
    #Campaign_dat <- paste('CAMPAIGN', Campaign_dat, sep='_')
    
    #read in input data
    #campId <- 
    # as.vector(read.csv(paste(dataPath, '\\test\\',
    #                         Campaign_dat, '_id',
    #                        Ext, sep=''))[1,1])
    
    CStart <-
        min(as.Date(as.character(read.csv(paste(dataPath, '\\Campaign_',
                                                campId, Ext, sep=''))$Engagement_Date),
                    format='%Y/%m/%d'))  
    
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
    matched_data2_temp <- function(xx){
        x <- xx[, -ncol(xx)]
        r <- 1:nrow(x)
        
        group <- ifelse(x[, match('cohort', var_list1)] ==1, "TEST", "CONTROL")
        temp1 <- lapply(1:Total_wks, function(i){
            gap <- i - x$engage_wk - 1
            pre_flag <- ifelse(gap > (-1-Pre_wks) & gap < 0, 1, ifelse(gap > 0, 0, -1))
            return(list(col_idx = 1:length(r), pre_flag, gap))
        })
        library(plyr)
        temp1_1 <- ldply(temp1, quickdf)
        temp1_1 <- arrange(temp1_1, col_idx, X2)
        
        x_t <- data.frame(t(x))
        test_fun <- function(x, y) {
            #row <- rownames(x_t) %in% var
            tmp <- x[y]
            return(tmp)
        }
        
        
        pre_flag <- tapply(temp1_1[, 2], temp1_1[, 1], function(x) x == 1)
        
        pre_othT_withMiss <- Map(test_fun, x_t[rownames(x_t) %in% othT, ], pre_flag)
        pre_othT_withMiss <- do.call(rbind.data.frame, pre_othT_withMiss)
        
        pre_nlT_withMiss <- Map(test_fun, x_t[rownames(x_t) %in% nlT, ], pre_flag)
        pre_nlT_withMiss <- do.call(rbind.data.frame, pre_nlT_withMiss)
        
        pre_othN_withMiss <- Map(test_fun, x_t[rownames(x_t) %in% othN, ], pre_flag)
        pre_othN_withMiss <- do.call(rbind.data.frame, pre_othN_withMiss)
        
        pre_nlN_withMiss <- Map(test_fun, x_t[rownames(x_t) %in% nlN, ], pre_flag)
        pre_nlN_withMiss <- do.call(rbind.data.frame, pre_nlN_withMiss)
        
        
        pre_period_othT <- rowSums(pre_othT_withMiss, na.rm = TRUE)
        pre_period_nlT <-  rowSums(pre_nlT_withMiss, na.rm = TRUE)
        pre_period_othN <- rowSums(pre_othN_withMiss, na.rm = TRUE)
        pre_period_nlN <- rowSums(pre_nlN_withMiss, na.rm = TRUE)
        
        pre_market_T <- pre_period_othT + pre_period_nlT
        pre_market_N <- pre_period_othN + pre_period_nlN
        
        pre_flag1 <- tapply(temp1_1[, 2], temp1_1[, 1], function(x) x == 0)
        
        post_othT_withMiss <- Map(test_fun, x_t[rownames(x_t) %in% othT, ], pre_flag1)
        post_othT_withMiss <- do.call(rbind.data.frame, post_othT_withMiss)
        
        post_nlT_withMiss <- Map(test_fun, x_t[rownames(x_t) %in% nlT, ], pre_flag1)
        post_nlT_withMiss <- do.call(rbind.data.frame, post_nlT_withMiss)
        
        post_othN_withMiss <- Map(test_fun, x_t[rownames(x_t) %in% othN, ], pre_flag1)
        post_othN_withMiss <- do.call(rbind.data.frame, post_othN_withMiss)
        
        post_nlN_withMiss <- Map(test_fun, x_t[rownames(x_t) %in% nlN, ], pre_flag1)
        post_nlN_withMiss <- do.call(rbind.data.frame, post_nlN_withMiss)
        
        post_period_othT <- rowSums(post_othT_withMiss, na.rm = TRUE)
        post_period_nlT <-  rowSums(post_nlT_withMiss, na.rm = TRUE)
        post_period_othN <- rowSums(post_othN_withMiss, na.rm = TRUE)
        post_period_nlN <- rowSums(post_nlN_withMiss, na.rm = TRUE)
        
        post_market_T <- post_period_othT + post_period_nlT
        post_market_N <- post_period_othN + post_period_nlN
        
        temp2 <- data.frame(pre_period_othT=pre_period_othT, 
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
        temp2_1=cbind(temp2[, 1:6] * Pre_wks/sapply(pre_flag, sum, na.rm=T),
                      temp2[, 7:12] * Pre_wks/sapply(pre_flag1, sum, na.rm=T))
        x1 <- x[,-match(c('engage_wk'), var_list1)]
        #     x2 <- as.vector(t(x1))
        #     names(x2) <- names(x1)
        gap <- temp1_1[seq(Total_wks, nrow(temp1_1), Total_wks), 3]
        return(data.frame(x1, group = group, gap = gap, temp2_1, engagement_date=xx$engagement_date))
    }
    
    
    
    start_time<- proc.time()
    start1 <- proc.time()
    num_pros <- n.cpu
    sfInit(parallel=TRUE, cpus=num_pros)
    sfLibrary("dplyr",character.only = TRUE)
    
    sfExport(
             'othT'
             ,'othN'
             ,'nlT'
             , 'nlN'
             , 'var_list1'
             , 'Total_wks'
             , 'Pre_wks')
    sfExport("ldply", namespace = "plyr")
    
    #run the function thru the combos
    #separate the whole matched_data1 into n.cpu parts
    set.seed(1)
    dt_flag <- sample(rep(1:n.cpu, length=nrow(matched_data1)))
    matched_data_sep <- lapply(1:n.cpu, function(i){
        matched_data1[dt_flag==i, ]
    })
    #matched_data2_1 <- sfClusterApplyLB(1:nrow(matched_data1), matched_data2_temp)
    matched_data2_1 <- sfClusterApplyLB(matched_data_sep, matched_data2_temp)
    sfStop()
    end_time<- proc.time()
    matched_data2 <- ldply(matched_data2_1, quickdf)
    cat('part1 used-', (proc.time()-start1)[3]/60, 'min!\n')
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
    anova_time0 <- proc.time()
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
    cat('anova time used-', (proc.time()-anova_time0)[3], 'sec!\n')
    
    # START: create points for graph
    # matched_data_forGp 
    
    get_graph_data_1 <- function(x){
        r <- 1:nrow(x)
        group <- ifelse(x$cohort==1, 'TEST', 'CONTROL')
        gap_vct <- as.data.frame(sapply(x$engage_wk, function(x)c(1:Total_wks) - x - 1)) #[128, 6] 
        gap_flag <- as.data.frame(
            sapply(as.data.frame(gap_vct), function(x)
                ifelse(x >(-1-Pre_wks) & x < (1+Pre_wks), T, F)
                )
            )
        breaks <- c(c(-Pre_wks+c(-Inf, -1, 4, 8, 13, 17, 21, 26)),
                    c(4, 8, 13, 17, 21, 26))
        Rel_month_vct <- c(-Inf, seq(-6, -1, 1), seq(1, 6, 1))
        bucket <- sapply(gap_vct, function(x)cut(x, breaks, right=F, labels=Rel_month_vct))
            
        rel_month <- lapply(r, function(i)bucket[gap_flag[, i], i])
        
        breaks1 <- c(-Inf, -Pre_wks-0.1, -14, -5, -1, Inf)
        bucket_period <- 
            lapply(r, function(i){
                tt <- cut(gap_vct[gap_flag[, i], i]
                    , breaks1
                    , right=T
                    , labels=c(-999, -3, -2, -1, 999)
                    )
                tt <- as.numeric(levels(tt)[tt])
                return(tt)
                })

            
        period <- lapply(r, function(i)
                ifelse(bucket_period[[i]]==999 | bucket_period[[i]]==-999, 
                       rel_month[[i]], bucket_period[[i]])
            )
        t1 <- proc.time()
        temp <- lapply(r, function(i){
            rel_trx <- as.numeric(x[i, match(nlT, var_list)][, gap_flag[, i]])
            rel_nrx <- as.numeric(x[i, match(nlN, var_list)][, gap_flag[, i]])
            rel_trx_other <- as.numeric(x[i, match(othT, var_list)][, gap_flag[, i]])
            rel_nrx_other <- as.numeric(x[i, match(othN, var_list)][, gap_flag[, i]])
            mkt_trx <- rel_trx + rel_trx_other
            mkt_nrx <- rel_nrx + rel_nrx_other
#             temp0 <- cbind(imsdr7=x$imsdr7[i],
#                           group=group[i], rel_wk=gap_vct[gap_flag[, i], i], 
#                           rel_month=rel_month[[i]], rel_trx=rel_trx, 
#                           rel_nrx=rel_nrx, rel_trx_other=rel_trx_other, 
#                           rel_nrx_other=rel_nrx_other, mkt_trx=mkt_trx, 
#                           mkt_nrx=mkt_nrx, period=period[[i]],
#                           spec=x$spec[i], zipcode=x$zipcode[i]
#                           )
            temp0 <- list(imsdr7=rep(x$imsdr7[i], length(rel_trx)),
                           group=rep(group[i], length(rel_trx)), rel_wk=gap_vct[gap_flag[, i], i], 
                           rel_month=rel_month[[i]], rel_trx=rel_trx, 
                           rel_nrx=rel_nrx, rel_trx_other=rel_trx_other, 
                           rel_nrx_other=rel_nrx_other, mkt_trx=mkt_trx, 
                           mkt_nrx=mkt_nrx, period=period[[i]],
                           spec=rep(x$spec[i], length(rel_trx)), zipcode=rep(x$zipcode[i], length(rel_trx))
            )
            return(temp0)
        })
        cat((proc.time()-t1)[3]/60, 'min!\n')
        
#         temp1 <- do.call(rbind.data.frame, temp)
        temp1 <- do.call(rbind, lapply(temp, data.frame, stringsAsFactors=F))
        # temp1 <- as_data_frame(temp)
#         temp1 <- ldply(temp, quickdf)
        #rel_trx <- mapply(function(a, b)a[b], as.matrix(x[, match(nlT, var_list)]), t(gap_flag))
        return(temp1)
        
    }
    start_time<- proc.time()
    start2 <- proc.time()
    num_pros <- n.cpu
    sfInit(parallel=TRUE, cpus=num_pros)
    
    sfExport('othT','othN','nlT', 'nlN'
             , 'var_list', 'Total_wks', 'Pre_wks'
             , 'Specfile', 'ZIP2Region', 'dataPath'
             , 'Ext')
    sfClusterEval(library("plyr"))
    sfClusterEval(library("dplyr"))
    
    #run the function thru the combos
    #separate the whole matched_data into n.cpu parts
    dt_flag <- sample(rep(1:n.cpu, length=nrow(matched_data)))
    matched_data_sep <- lapply(1:n.cpu, function(i){
        matched_data[dt_flag==i, ]
    })
    
    graph_data2_temp <- sfClusterApplyLB(matched_data_sep, get_graph_data_1)
    #graph_data2_temp <- sfClusterApplyLB(1:2000, get_graph_data_1)
    sfStop()
    end_time<- proc.time()
    
    check_dim <- ldply(lapply(graph_data2_temp, 
                              function(i){
                                  return(dim(i))
                              }), 
                       quickdf)
    
    graph_data_2 <- ldply(graph_data2_temp, quickdf)
    
    #create 'speciatly' and 'region' if 'Specfile' and 'ZIP2Region' files are given.
    
    if(Specfile != ''){
        spec_dic <- 
            read.table(paste(dataPath, '\\', Specfile, Ext, sep=''), sep=',', header=T
                       , stringsAsFactors=F)
        idx_spec <- match(graph_data_2$spec, spec_dic$SPEC_CD)
        graph_data_2$specialty <-
            ifelse(!is.na(idx_spec), 
                   as.character(spec_dic[idx_spec,
                                         'CLASS_DESC']), 'ALL OTHER')
    }else{
        graph_data_2$specialty <- "ALL SPECIALTY"
    }
    if(ZIP2Region != ''){
        reg_dic <- 
            read.table(paste(dataPath, '\\', ZIP2Region, Ext, sep=''), sep=',', header=T
                       , stringsAsFactors=F)
        graph_data_2$region <-
            ifelse(!is.na(match(graph_data_2$zipcode, reg_dic$ZIPCODE)), 
                   as.character(reg_dic[match(graph_data_2$zipcode, reg_dic$ZIPCODE),
                                        "REGION"]), 'OTHER')
    }else{
        graph_data_2$region <- 'ALL REGION'
    }   
    
    graph_data_2$spec <- NULL
    graph_data_2$zipcode <- NULL
    var_list2 <- names(graph_data_2)
    cat('part2 used-', (proc.time()-start2)[3]/60, 'min!\n')
    
    

    #    ***********  Get Monthly summary ****************;
    target_trx <- grep('trx|nrx', var_list2, value=T)
    graph_data_summary <- graph_data_2[, c(target_trx, 'group', 'rel_month')] %>%
        group_by(group, rel_month) %>%
        summarise_each(funs(sum))
    
    hcp_cnt <- graph_data_2[, c('imsdr7', 'group', 'rel_month')] %>%
        group_by(group, rel_month) %>%
        summarise_each(funs(length(unique(.))))
    
    
    
    graph_data_summary_1 <- 
        left_join(graph_data_summary[order(graph_data_summary$group), ],
                  hcp_cnt, by=c("group" = 'group', "rel_month" = "rel_month"))
    names(graph_data_summary_1)[length(graph_data_summary_1)] <- 'hcp_cnt'
    
    #*******************  Get Post Match test/control comparison ****************;
    monthly_data <- cbind(graph_data_2[, c(target_trx
                                           , 'specialty'
                                           , 'region'
                                           , 'group'
                                           , 'imsdr7'
                                           , 'period')]
                          , freq=rep(1, nrow(graph_data_2))) %>%
        group_by(specialty, region, group, imsdr7, period) %>%
        summarise_each(funs(sum(., na.rm=T))) %>%
        filter(as.numeric(period) < 0)
    
    
    
    
    monthly_mean <- monthly_data[, c(target_trx
                                           , 'specialty'
                                           , 'region'
                                           , 'group'
                                           # , 'imsdr7'
                                           , 'period')] %>%
        group_by(specialty, region, group,  period) %>%
        summarise_each(funs(mean(., na.rm=T)))
    
    monthly_std <- monthly_data[, c(target_trx
                                        , 'specialty'
                                        , 'region'
                                        , 'group'
                                        # , 'imsdr7'
                                        , 'period')] %>%
        group_by(specialty, region, group,  period) %>%
        summarise_each(funs(sd(., na.rm=T)))
    
    
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

    monthly_all_region_mean <- monthly_data[, c(target_trx
                                                , 'specialty'
                                                , 'group'
                                                , 'period'
                                                )] %>%
        group_by(specialty, group, period) %>%
        summarise_each(funs(mean(., na.rm=T)))
    
    monthly_all_region_std <- monthly_data[, c(target_trx
                                               , 'specialty'
                                               , 'group'
                                               , 'period'
                                                )] %>%
        group_by(specialty, group, period) %>%
        summarise_each(funs(sd(., na.rm=T)))
    
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
        levels <- levels(as.factor(input$group))
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
                               CampaignId=139, 
                               Specfile='spec_grp', 
                               ZIP2Region='ZIP_region', 
                               n.cpu=8))


#esti_df <- part2_result$esti_df
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

