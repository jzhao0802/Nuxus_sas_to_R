#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#
#   Project: NEXXUS
#   Purpose: Matching Process
#   Version: 0.1
#   Programmer: Xin Huang
#   Date: 09/09/2015
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

# read in the fake data for testing
# hcps_rx_4match <- read.csv("hcps_rx_4match.csv", header = TRUE,
#                            stringsAsFactors = FALSE)

ps_match <- function(data, group, id, mvars, wts, dmaxk, dmax = NULL, dist = 1,
                     ncontls = 1, time = NULL, transf = 0, seedca, seedco,
                     rsp_var, rsp_ind, pscore_cut = 0.0001){
  ##@@parameters:
    # DATA    :  data set containing cases and potential controls
    # GROUP   :  variable defining cases. Group:1 if case 0 if control
    # ID      :  CHARACTER ID variable for the cases and controls
    # MVARS   :  List of numeric matching variables common to both case and control
    # WTS     :  List of non-negative weights corresponding to each matching vars
    # DMAXK   :  List of non-negative values corresponding to each matching vars
    # DMAX    :  Largest value of Distance(Dij) to be considered as a valid match
    # DIST    :  Type of distance to use 1:absolute difference; 2:Euclidean
    # NCONTLS :  The number of controls to match to each case 
    # TIME    :  Time variable used for risk set matching, only if ctrl time > case being matched 
    # TRANSF  :  Whether all matching vars are to be transformed  (0:no, 1:standardize, 2:use ranks) 
    # SEEDCA  :  Seed value used to randomly sort the cases prior to match 
    # SEEDCO  :  Seed value used to randomly sort the controls prior to match 
    # Rsp_var :  To include propensity score in match. Response variable to calculate propensity score  
    # Rsp_Ind :  Independent variable list to be used for propensity score calculation, dafault:&MVars 
    # Pscore_cut :  set propensity score diff < 0.005 as valid match 
  ##@@return
    # a list  :  contains the matched result, unmatched case and unmatched controls.
  
#   data = match_raw_data
#   group = "case"
#   id = "id"
#   mvars = c("sex", "age")
#   wts = c(1, 1)
#   dmaxk = c(0, 2)
#   #dmax = 
#   dist = 1
#   ncontls = 2
#   #time
#   transf = 0
#   seedca = 234
#   seedco = 489
#   print = TRUE
#   rsp_var = "resp"
#   rsp_ind = c("sex", "age", "income")
#   pscore_cut = 0.0001
#   data    = hcps_rx_4match
#   group   = "cohort"
#   id       = "imsdr7"
#   mvars   = c("spec_grp", "reg_num", "lat_nbr", "lon_nbr",
#               "pre_nov12", "pre_nov", "pscore")
#   wts	   =  c(1, 1, 1, 1, 5, 5, 5)
#   dmaxk   = c(0, 0, 1, 1, 0.001453205, 0.002909882, 0.001)
#   transf  = 0
#   dist    = 1
#   ncontls = 1
#   seedca  = 279
#   seedco  = 389
#   print   = FALSE
  
  require(dplyr)
  require(doParallel)
  # do the logistic regression to get the propensity
  if (! missing(rsp_var)) {
    if (missing(rsp_ind)) {
      rsp_ind <- mvars
    }
    formu <- paste(rsp_var, " ~ ", paste(rsp_ind, collapse = " + "))
    glm_model <- glm(formu, data, family = "binomial")
    p_score <- predict(glm_model, data, type = "response")
    .ndata <- cbind(data, pscore = p_score)
    
    data <- .ndata
    mvars <- c(mvars, "pscore")
    wts <- c(wts, 1)
    if (! missing(dmaxk)) dmaxk <- c(dmaxk, pscore_cut)
  }
  
  bad <- 0
  if (missing(data)) {
    bad <- 1
    stop("ERROR: NO DATASET SUPPLIED") 
  }
  
  if (missing(id)) {
    bad <- 1
    stop("ERROR: NO ID VARIABLE SUPPLIED")  
  }
  
  if (missing(group)) {
    bad <- 1
    stop("ERROR: NO CASE(1)/CONTROL(0) GROUP VARIABLE SUPPLIED")
  }
  
  if (missing(wts)) {
    bad <- 1  
    stop("ERROR: NO WEIGHTS SUPPLIED")
  }
  
  nvar <- length(mvars)
  nwts <- length(wts)
  
  if(nvar != nwts) {
    bad <- 1
    stop("ERROR: #VARS MUST EQUAL #WTS")
  }
  
  nk <- length(dmaxk)
  if (nk > nvar) nk <- nvar
   
  v <- mvars
  w <- nwts
  if(any(w < 0)) {
    bad = 1
    stop("EERROR: WEIGHTS MUST BE NON-NEGATIVE")
  }
  
  if (nk > 0) {
    k <- dmaxk
    if(any(k < 0)){
      bad = 1
      stop("ERROR: DMAXK VALUES MUST BE NON-NEGATIVE")
    }
  }
  
  #### for matching #####
  
  
  # remove the rows with missing values
  .check <- data
  .check[ ".id"] <- data[, id]
  .check <- .check[complete.cases(.check), ]
  
  # standardize the vars
  if (transf == 1) {
    .stdzd <- scale(.check[, mvars], center = TRUE, scale = TRUE)
    .caco <- cbind(.check[, setdiff(colnames(.check), mvars)], .stdzd)
  } else {
    .caco <- .check
  }
  
  # for case dataset
  .case <- .caco[.caco[, group] == 1, ]
  .case[, ".idca"] <- .case[, id]
  
  if (!missing(time)) .case[, ".catime"] <- .case[, time]
  
  tmp <- paste(".ca", 1:nvar, sep = "") 
  .case[, tmp] <- .case[, v]
  
  if (! missing(seedca)) {
    set.seed(seedca)
    .case[, ".r"] <- runif(nrow(.case))
  } else {
    .case[, ".r"] <- 1
  }
  
  if (missing(time)) {
    .case <- .case[, c('.idca', tmp, ".r", mvars)]
  } else {
    .case <- .case[, c('.idca', tmp, ".r", mvars, ".catime")]
  }
  
  nca <- nrow(.case)
  
  # for control dataset
  
  .cont <- .caco[.caco[, group] == 0, ]
  
  .cont[, ".idco"] <- .cont[, id]
  if (!missing(time)) .cont[, ".cotime"] <- .cont[, time]
  
  tmp1 <- paste(".co", 1:nvar, sep = "")
  .cont[, tmp1] <- .cont[, v]
  
  if (! missing(seedco)) {
    set.seed(seedco)
    .cont[, ".r"] <- runif(nrow(.cont))
  } else {
    .cont[, ".r"] <- 1
  }
  
  if (missing(time)) {
    .cont <- .cont[, c('.idco', tmp1, ".r", mvars)]
  } else {
    .cont <- .cont[, c('.idco', tmp1, ".r", mvars, ".cotime")]
  }
  
  nco <- nrow(.cont)
  
  bad2 <- 0
  
  if (nco < nca * ncontls) {
    bad2 <- 1
    stop("ERROR: NOT ENOUGH CONTROLS TO MAKE REQUESTED MATCHES")
  }
  
  .cont <- arrange(.cont, .r, .idco)
  
  #### do the matching #####
  dmax_missing <- is.null(dmax)
  time_missing <- is.null(time)
  
  eq_vars_flag <- which(dmaxk == 0)
  if (length(eq_vars_flag)) {
    split <- unique(.case[, mvars[eq_vars_flag]])
    split_critirea <- Map(function(x, y) paste(x, "==", y, sep = " "),
                          mvars[eq_vars_flag],
                          split)
    
    split_critirea_m <- data.frame(split_critirea)
    split_critirea_m1 <- NULL
    for (i in 1:ncol(split_critirea_m)) {
      if (i == 1){
        split_critirea_m1 <- 
          paste(split_critirea_m[, i], split_critirea_m1, sep = "")
      } else {
        split_critirea_m1 <- 
          paste(split_critirea_m[, i], "&", split_critirea_m1, sep = " ")
      }
    } 
    
    writeLines(c(" "), "log.txt")
    cl <- makeCluster(floor(detectCores() / 4))
    registerDoParallel(cl)
    match_results <- 
      foreach (i = 1:length(split_critirea_m1),  
               .packages = "dplyr") %dopar% {                       
    # distance matrix
      tryCatch(
      {.case1 <- subset(.case, eval(parse(text = split_critirea_m1[i])))
       .cont1 <- subset(.cont, eval(parse(text = split_critirea_m1[i])))
       d.substract <- apply(.case1[, tmp], 
                            1, 
                            function(x) apply(.cont1[, tmp1], 1, function(y) x - y))
       d.substract <- as.vector(d.substract)
       d.matrix <- matrix(d.substract, ncol = length(tmp), byrow = TRUE)
       rm(d.substract)
       gc()
       gap_flag <- apply(d.matrix, 1, function(x) all(abs(x) <= dmaxk)) 
       gap_flag_matrix <- matrix(gap_flag, nrow = nrow(.case1), byrow = TRUE)
       rm(gap_flag)
       gc()
       
       if (dist == 2) {
         distance <- apply(d.matrix, 1, function(x) sqrt(sum((x * wts) ^ 2)))
       } 
       if (dist == 1) {
         distance <- apply(d.matrix, 1, function(x) sum((abs(x) * wts)))
       }
       distance_matrix <- matrix(distance, nrow = nrow(.case1), byrow = TRUE)
       rm(distance)
       gc()
       
       if (! time_missing) {
         time_substract <- 
           as.vector(sapply(.case1[, ".catime"], function(x) .cont1[, ".cotime"] > x))
         time_flag <- matrix(time_subtract, nrow = nrow(.case1), byrow = TRUE)
       }
       
       # use this matrix to determine the matched control 
       if (time_missing) {
         location_flag <- gap_flag_matrix
       } else {
         location_array <- array(nca, nco, 2)
         location_array[, , 1] <- gap_flag_matrix
         location_array[, , 2] <- time_flag
         location_flag <- apply(location_array, 3, all)  
       }
       
       matched_matrix <- matrix(0, nrow = nrow(.case1), ncol = ncontls)
       d_matrix <- matrix(0, nrow = nrow(.case1), ncol = ncontls)
       matched_flag <- rep(0, nrow(.cont1))
       for (i in 1:ncontls) {
         for (j in 1:nrow(.case1)) {
           idx <- which(location_flag[j, ])
           if (!length(idx)) {
             matched_matrix[j, i] <- NA
             next
           } else {
             while(TRUE) {
               matched <- idx[which.min(distance_matrix[j, idx])[1]]  
               if (time_missing) {
                 if (matched_flag[matched]) {
                   idx <- setdiff(idx, matched)
                   if (length(idx) == 0) {
                     matched_matrix[j, i] <- NA
                     break
                   }
                   next
                 } else {
                   matched_matrix[j, i] <- matched 
                   d_matrix[j, i] <- distance_matrix[i, matched]
                   matched_flag[matched] <- 1
                   break
                 }
               } else {
                 if (matched_flag[matched] |
                       distance_matrix[j, matched] > dmax ) {
                   idx <- setdiff(idx, matched)
                   if (length(idx) == 0) {
                     matched_matrix[j, i] <- NA
                     break
                   }
                   next
                 } else {
                   matched_matrix[j, i] <- matched 
                   d_matrix[j, i] <- distance_matrix[j, matched]
                   matched_flag[matched] <- 1
                   break
                 }
               }
             }    
           }    
         }
       } 
       
       matched_matrix <- data.frame(matched_matrix)
       
       #get the match and unmatched case and control
       .case1_matched_flag <- lapply(matched_matrix, 
                                     function(x) (1:nrow(.case1))[!is.na(x)]) 
       .case1_matched <- lapply(.case1_matched_flag, function(x) .case1[x,])
       
       .case1_unmatched_flag <- apply(matched_matrix, 1, function(x) all(is.na(x))) 
       .case1_unmatched <- .case1[.case1_unmatched_flag, ]
       
       .cont1_matched_flag <- lapply(matched_matrix, function(x) x[!is.na(x)]) 
       .cont1_matched <- lapply(.cont1_matched_flag, function(x) .cont1[x, ])
       
       .cont1_unmatched_flag <- setdiff(1:nrow(.cont1), unlist(.cont1_matched_flag))
       .cont1_unmatched <- .cont1[.cont1_unmatched_flag, ]
       
       d <- Map(function(x, y) {
         if (!length(x) & !length(y)) {
           NULL
         } else {
           d.tmp <- matrix(distance_matrix[x, y], length(x), 1)
           colnames(d.tmp) <- "distance"
           d.tmp
         }
       }, .case1_matched_flag, .cont1_matched_flag)
       gap <- lapply(.case1_matched_flag, 
                     function(x) {
                       if (length(x) > 0){
                         d.tmp <- d.matrix[(x - 1) * nrow(.cont1) + x, ]
                         if (is.null(dim(d.tmp))) {
                           d.tmp <- matrix(d.tmp, nrow = 1, byrow = TRUE)
                         } 
                         colnames(d.tmp) <- paste(mvars, "_diff", sep = "")
                         d.tmp
                       } else {
                         NULL
                       }
                     })
       
       final_matched <- lapply(1:length(.case1_matched), 
                               function(x){
                                 cbind(.case1_matched[[x]], 
                                       .cont1_matched[[x]],
                                       d[[x]],
                                       gap[[x]])
                               })
       
       if (length(.case1_matched) > 0) {
         final_case <- lapply(1:length(.case1_matched), 
                              function(x){
                                case.tmp <- .case1_matched[[x]][, ".idca"]
                                case.tmp <- data.frame(case.tmp)
                                colnames(case.tmp) <- id
                                case.tmp <- inner_join(data, case.tmp, by = id)
                                case.tmp <- data.frame(case.tmp, 
                                                       cohort = rep("ca", nrow(case.tmp)))
                                case.tmp 
                              }) 
         
         final_cont <- lapply(1:length(.cont1_matched), 
                              function(x){
                                cont.tmp <- .cont1_matched[[x]][, ".idco"]
                                cont.tmp <- data.frame(cont.tmp)
                                colnames(cont.tmp) <- id
                                cont.tmp <- inner_join(data, cont.tmp, by = id)
                                cont.tmp <- data.frame(cont.tmp, 
                                                       cohort = rep("co", nrow(cont.tmp)))
                                cont.tmp 
                              })
         final_matched_output <- do.call("rbind",
                                         list(final_case[[1]], 
                                              do.call("rbind", final_cont)))
       } else {
         final_case <- NULL
         final_cont <- NULL
         final_matched_output <- NULL
       }
       list(final_matched, final_matched_output, .case1_unmatched, .cont1_unmatched)},
      error = function(e) {
        cat(paste("The error:", e, "happened in", i, "cohort"),
            file = "log.txt", 
            append = TRUE)})
      }
  stopCluster(cl)
    
  final_matched <- do.call("rbind", lapply(match_results, function(x) x[[1]][[1]]))
  final_matched_output <- do.call("rbind", lapply(match_results, function(x) x[[2]]))
  .case_unmatched <- do.call("rbind", lapply(match_results, function(x) x[[3]]))
  .cont_unmatched <- do.call("rbind", lapply(match_results, function(x) x[[4]]))
  .cont_unmatched_flag <- unique(.cont_unmatched[, ".idco"])
  .cont_matched_flag <- unique(subset(final_matched_output, cohort == 0)[, id])
  .cont_unmatched_other <- 
    subset(.cont, ! .idco %in% c(.cont_unmatched_flag, .cont_matched_flag))
  .cont_unmatched <- rbind(.cont_unmatched, .cont_unmatched_other)
  
#   if (print) {
#     print("******* The Summary of Matched Cases & Matched Control: ********")
#     print(apply(final_matched, 2, fivenum))
#     
#     print("******* The Summary of Matched Cases ********")
#     print(apply(subset(final_matched_output, cohort == 1), 2, fivenum))
#     
#     print("******* The Summary of Unmatched Cases ********")
#     print(apply(.case_unmatched, 2, fivenum))
#     
#     print("******* The Summary of Matched Controls ********")
#     print(apply(subset(final_matched_output, cohort == 0), 2, fivenum))
#     
#     print("******* The Summary of Unmatched Controls ********")
#     print(apply(.cont1_unmatched, 2, fivenum))
#  }
  } else {
    # distance matrix
    d.substract <- apply(.case[, tmp], 
                         1, 
                         function(x) apply(.cont[, tmp1], 1, function(y) x - y))
    
    d.substract <- as.vector(d.substract)
    d.matrix <- matrix(d.substract, ncol = length(tmp), byrow = TRUE)
    rm(d.substract)
    gc()
    gap_flag <- apply(d.matrix, 1, function(x) all(abs(x) <= dmaxk)) 
    gap_flag_matrix <- matrix(gap_flag, nrow = nca, byrow = TRUE)
    rm(gap_flag)
    gc()
    
    if (dist == 2) {
      distance <- apply(d.matrix, 1, function(x) sqrt(sum((x * wts) ^ 2)))
    } 
    if (dist == 1) {
      distance <- apply(d.matrix, 1, function(x) sum((abs(x) * wts)))
    }
    distance_matrix <- matrix(distance, nrow = nca, byrow = TRUE)
    rm(distance)
    gc()
    
    if (!time_missing) {
      time_substract <- 
        as.vector(sapply(.case[, ".catime"], function(x) .cont[, ".cotime"] > x))
      time_flag <- matrix(time_subtract, nrow = nca, byrow = TRUE)
    }
    
    # use this matrix to determine the matched control 
    if (time_missing) {
      location_flag <- gap_flag_matrix
    } else {
      location_array <- array(nca, nco, 2)
      location_array[, , 1] <- gap_flag_matrix
      location_array[, , 2] <- time_flag
      location_flag <- apply(location_array, 3, all)  
    }
    
    matched_matrix <- matrix(0, nrow = nca, ncol = ncontls)
    d_matrix <- matrix(0, nrow = nca, ncol = ncontls)
    matched_flag <- rep(0, nco)
    for (i in 1:ncontls) {
      for (j in 1:nca) {
        idx <- which(location_flag[j, ])
        if (!length(idx)) {
          matched_matrix[j, i] <- NA
          next
        } else {
          while(TRUE) {
            matched <- idx[which.min(distance_matrix[j, idx])[1]]  
            if (dmax_missing) {
              if (matched_flag[matched]) {
                idx <- setdiff(idx, matched)
                if (length(idx) == 0) {
                  matched_matrix[j, i] <- NA
                  break
                }
                next
              } else {
                matched_matrix[j, i] <- matched 
                d_matrix[j, i] <- distance_matrix[i, matched]
                matched_flag[matched] <- 1
                break
              }
            } else {
              if (matched_flag[matched] |
                    distance_matrix[j, matched] > dmax ) {
                idx <- setdiff(idx, matched)
                if (length(idx) == 0) {
                  matched_matrix[j, i] <- NA
                  break
                }
                next
              } else {
                matched_matrix[j, i] <- matched 
                d_matrix[j, i] <- distance_matrix[j, matched]
                matched_flag[matched] <- 1
                break
              }
            }
          }    
        }    
      }
    } 
    
    matched_matrix <- data.frame(matched_matrix)
    
    .case_matched_flag <- lapply(matched_matrix, function(x) (1:nca)[!is.na(x)]) 
    .case_matched <- lapply(.case_matched_flag, function(x) .case[x,])
    
    .case_unmatched_flag <- apply(matched_matrix, 1, function(x) all(is.na(x))) 
    .case_unmatched <- .case[.case_unmatched_flag, ]
    
    .cont_matched_flag <- lapply(matched_matrix, function(x) x[!is.na(x)]) 
    .cont_matched <- lapply(.cont_matched_flag, function(x) .cont[x, ])
    
    .cont_unmatched_flag <- setdiff(1:nco, unlist(.cont_matched_flag))
    .cont_unmatched <- .cont[.cont_unmatched_flag, ]
    
    d <- Map(function(x, y) {
      if (!length(x) & !length(y)) {
        NULL
      } else {
        d.tmp <- matrix(distance_matrix[x, y], length(x), 1)
        colnames(d.tmp) <- "distance"
        d.tmp
      }
    }, .case_matched_flag, .cont_matched_flag)
    gap <- lapply(.case_matched_flag, 
                  function(x) {
                    if (length(x) > 0){
                      d.tmp <- d.matrix[(x - 1) * nco + x, ]
                      colnames(d.tmp) <- paste(mvars, "_diff", sep = "")
                      d.tmp
                    } else {
                      NULL
                    }
                  })
    
    final_matched <- lapply(1:length(.case_matched), 
                            function(x){
                              cbind(.case_matched[[x]], 
                                    .cont_matched[[x]],
                                    d[[x]],
                                    gap[[x]])
                            })
    
    final_case <- lapply(1:length(.case_matched), 
                         function(x){
                           force(x)
                           tmp <- .case_matched[[x]][, ".idca"]
                           tmp <- data.frame(tmp)
                           colnames(tmp) <- id
                           tmp <- inner_join(data, tmp, by = id)
                           tmp <- data.frame(tmp, cohort = rep("ca", nrow(tmp)))
                           tmp
                         }) 
    
    final_cont <- lapply(1:length(.cont_matched), 
                         function(x){
                           force(x)
                           tmp <- .cont_matched[[x]][, ".idco"]
                           tmp <- data.frame(tmp)
                           colnames(tmp) <- id
                           tmp <- inner_join(data, tmp, by = id)
                           tmp <- data.frame(tmp, cohort = rep("co", nrow(tmp)))
                           tmp
                         })
    
    final_matched_output <- do.call("rbind",
                                    list(final_case[[1]], 
                                         do.call("rbind", final_cont)))
    
#     if (print) {
#       print("******* The Summary of Matched Cases & Matched Control: ********")
#       print(lapply(final_matched, function(x) apply(x, 2, fivenum)))
#       
#       print("******* The Summary of Matched Cases ********")
#       print(lapply(.case_matched, function(x) apply(x, 2, fivenum)))
#       
#       print("******* The Summary of Unmatched Cases ********")
#       print(apply(.case_unmatched, 2, fivenum))
#       
#       print("******* The Summary of Matched Controls ********")
#       print(lapply(.cont_matched, function(x) apply(x, 2, fivenum)))
#       
#       print("******* The Summary of Unmatched Controls ********")
#       print(apply(.cont_unmatched, 2, fivenum))
#     }
    del_objs <- setdiff(ls(), 
                        c("final_matched_output", ".case_unmatched", ".cont_unmatched"))
    rm(list = del_objs)
    gc()
  }
  return(list(matched = final_matched_output, 
              unmatched_case = .case_unmatched,
              unmatched_control = .cont_unmatched))
}

# user  system elapsed 
# 13.02   43.53  747.21
# system.time(ps_result <- ps_match(data = hcps_rx_4match,
#                                   group = "cohort",
#                                   id = "imsdr7",
#                                   mvars = c("spec_grp", "reg_num", 
#                                             "lat_nbr", "lon_nbr",
#                                             "pre_nov12", "pre_nov", "pscore"),
#                                   wts =  c(1, 1, 1, 1, 5, 5, 5),
#                                   dmaxk = c(0, 0, 1, 1, 0.001453205,
#                                             0.002909882, 0.001),
#                                   transf = 0,
#                                   dist = 1,
#                                   ncontls = 1,
#                                   seedca = 279,
#                                   seedco = 389))
