library(tidyverse)
library(ROI)
library(ROI.plugin.glpk)
library(ompr)
library(ompr.roi)
library(readr)
library(openxlsx)
library(Surrogate)



random_generator <- function(lots, bidders, lotthreshold){
    
    nooflots = lots
    # Total bidders applied are 
    noofbidders = bidders
    
    lotslist = seq(1:nooflots)
    bidderslist = seq(1:noofbidders)
    
    
    # lot_caseloads = c(15, 14,	1,	3,	5,	9,	1,	13,	7,	2,	10,	9)
    rand_vec <- RandVec(a=2, b=10, s=100, n=nooflots, m=1)
    lot_caseloads <- round(as.numeric(unlist(rand_vec)))
    lot_capital = sample(1:30, size =nooflots, replace = T)
    lot_capital_pc = sample(max(lot_capital):50, size =nooflots, replace = T)
    bidder_capital = sample(max(lot_capital):60, size =noofbidders, replace = T)
    bidder_capital_pc = sample(max(bidder_capital):100, size =noofbidders, replace = T)
    
    lot_case_load_threshold = lotthreshold
    
    pqp_score = matrix(data =c(floor(runif(nooflots * noofbidders,min=10, max=100))), nrow = noofbidders,
                       ncol = nooflots)
    
    model <- MIPModel() %>% 
        add_variable(x[i, j], i = 1:noofbidders, j = 1:nooflots, type = "binary") %>% 
        add_constraint(sum_expr(x[i, j], i = 1:noofbidders) == 1 , j = 1:nooflots) %>% 
        add_constraint(sum_expr(lot_caseloads[j] * x[i ,j], j = 1:nooflots) <= lot_case_load_threshold,
                       i = 1:noofbidders) %>% 
        add_constraint(sum_expr(lot_capital[j] * x[i ,j], j = 1:nooflots) <= bidder_capital[i],
                       i = 1:noofbidders) %>% 
        add_constraint(sum_expr(lot_capital_pc[j] * x[i ,j], j = 1:nooflots) <= bidder_capital_pc[i],
                       i = 1:noofbidders) %>% 
        set_objective(sum_expr(pqp_score[i, j] * x[i ,j],  i = 1:noofbidders,
                               j = 1:nooflots) , "min") 
    
    
    
    
    result <- solve_model(model, with_ROI(solver = "glpk", verbose = TRUE))
    minimumpqp <- result %>% 
                    objective_value()
    status <- result %>% 
                    solver_status()
    
    if (result %>% solver_status()=="optimal") {
        
        
        
        matching <- result %>% 
            get_solution(x[i , j]) %>% 
            filter(value > 0) %>%
            mutate(Bidder = bidderslist[i],Lots = j
                   , Lot_Capital = lot_capital[j],
                   Bidder_cap = bidder_capital[i],
                   Lot_Capital_PC = lot_capital_pc[j],
                   Bidder_cap_PC = bidder_capital_pc[i],
                   Lot_CaseLoad = lot_caseloads[j])%>%
            dplyr::select(Lots,Bidder, Lot_Capital, Bidder_cap,
                   Lot_Capital_PC,Bidder_cap_PC,Lot_CaseLoad,i,j) %>% 
            rowwise() %>% 
            mutate(pqpscore = pqp_score[i,j]) %>% 
            dplyr::select(Lots,Bidder,Lot_CaseLoad,pqpscore, Lot_Capital, Bidder_cap,
                   Lot_Capital_PC,Bidder_cap_PC)
        
        return(list(minimumpqp,status,matching ))
        
    } else {
        print("No optimal solution found, Make sure Total number bidders accounts at least 60% of total Lots")
    }
    
    
    
    
    
}




excel_input_generator <- function(excelfile, lotthreshold){
    
    ##import data from excel template
    
    
    lotsdata_excel <- openxlsx::read.xlsx(excelfile, sheet ='lotsdata')
    pqp_score_excel <- openxlsx::read.xlsx(excelfile, sheet ='pqpscore')
    biddersdata_excel <- openxlsx::read.xlsx(excelfile, sheet ='bidderdata')
    pqp_score_excel <- pqp_score_excel[order(pqp_score_excel$LotID, pqp_score_excel$BidderID),]
    
    bidders <-  sort(c(unique(pqp_score_excel$BidderID)))
    lots <-   sort(c(unique(pqp_score_excel$LotID)))
    
    noofbidders <- length(bidders)
    nooflots <- length(lots)
    
    
    lot_caseloads = c(lotsdata_excel$CaseLoads)*100
    
    lot_capital = c()
    lot_capital_pc = c()
    for (i in lots){
        
        lot_capital <- c(lot_capital,pull(subset(lotsdata_excel, LotID == i)['LotCapital']))
        lot_capital_pc <- c(lot_capital_pc,pull(subset(lotsdata_excel, LotID == i)['LotCapitalPC']))
        
    }
    
    
    bidder_capital = c()
    bidder_capital_pc = c()
    
    for (i in bidders){
        
        bidder_capital <- c(bidder_capital,pull(subset(biddersdata_excel, BidderID == i)['BidderCapital']))
        bidder_capital_pc <- c(bidder_capital_pc,pull(subset(biddersdata_excel, BidderID == i)['BidderCapitalPC']))
        
    }
    
    
    lot_case_load_threshold = lotthreshold
    
    
    pqp_score = as.matrix(spread(subset(pqp_score_excel
                                        , select = c(BidderID, LotID,PQPScore)),
                                 LotID, PQPScore)[,-1])
    
    
    pqp_score[is.na(pqp_score)] <- 1000000
    pqp_score
    
    model <- MIPModel() %>% 
        add_variable(x[i, j], i = 1:noofbidders, j = 1:nooflots, type = "binary") %>% 
        add_constraint(sum_expr(x[i, j], i = 1:noofbidders) == 1 , j = 1:nooflots) %>% 
        add_constraint(sum_expr(lot_caseloads[j] * x[i ,j], j = 1:nooflots) <= lot_case_load_threshold,
                       i = 1:noofbidders) %>% 
        add_constraint(sum_expr(lot_capital[j] * x[i ,j], j = 1:nooflots) <= bidder_capital[i],
                       i = 1:noofbidders) %>% 
        add_constraint(sum_expr(lot_capital_pc[j] * x[i ,j], j = 1:nooflots) <= bidder_capital_pc[i],
                       i = 1:noofbidders) %>% 
        set_objective(sum_expr(pqp_score[i, j] * x[i ,j],  i = 1:noofbidders,
                               j = 1:nooflots) , "min") 
    
    
    
    result <- solve_model(model, with_ROI(solver = "glpk", verbose = TRUE))
    
    minimumpqp <- result %>% 
                 objective_value()
    status <- result %>% 
                    solver_status()
    
    matching <- result %>% 
        get_solution(x[i , j]) %>% 
        filter(value > 0) %>% 
        mutate(Bidder = bidders[i],Lots = lots[j],
               Lot_Capital = lot_capital[j],
               Bidder_cap = bidder_capital[i],
               Lot_Capital_PC = lot_capital_pc[j],
               Bidder_cap_PC = bidder_capital_pc[i],
               Lot_CaseLoad = lot_caseloads[j]) %>%
        select(Lots,Bidder, Lot_Capital, Bidder_cap,
               Lot_Capital_PC,Bidder_cap_PC,Lot_CaseLoad,i,j) %>%
        rowwise() %>%
        mutate(pqpscore = pqp_score[i,j]) %>%
        select(Lots,Bidder,Lot_CaseLoad,pqpscore, Lot_Capital, Bidder_cap,
               Lot_Capital_PC,Bidder_cap_PC)
    
    return(list(minimumpqp,status,matching ))
    
    
}


