aggregated <- graph_data_3[,target_trx] %>%
      
      group_by(group, rel_month) %>%
      summarize(trx = sum(trx), nrx = sum(nrx))


library(reshape)
summary(Indometh)
wide <- reshape(Indometh
                , v.names = "conc"
                , idvar = "Subject"
                , timevar = "time"
                , direction = "wide")
wide


state.x77 <- as.data.frame(state.x77)
state.x77$state <- rownames(state.x77)
long <- reshape(state.x77, idvar = c("state", 'whether')
                # , ids = row.names(state.x77)
                , times = setdiff(names(state.x77), c('state', 'whether'))
                , timevar = "Characteristic"
                , v.names = 'Value'
                , varying = list(setdiff(names(state.x77), c('state', 'whether')))
                , direction = "long")

wide <- reshape(long
                , v.names='Value'
                , idvar = c('whether', 'state')
                , timevar = 'Characteristic'
                , direction = 'wide'
                
                )