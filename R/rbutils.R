#Utility functions for meta-analysis

#load packages required by functions in this file
library(dplyr, quietly=TRUE)
library(stringr, quietly=TRUE)
library(readxl, quietly=TRUE)
library(xlsx, quietly=TRUE)
library(tools, quietly=TRUE)
library(meta, quietly=TRUE)
library(ggmcmc, quietly=TRUE)

############################
#Functions
factorToCharacter = function(df){
  #takes a data frame and converts all factor variables to characters
  df[sapply(df, is.factor)] <- lapply(df[sapply(df, is.factor)], 
                                      as.character)
  df
}

capwords <- function(s, strict = FALSE) {
  cap <- function(s) paste(toupper(substring(s, 1, 1)),
{s <- substring(s, 2); if(strict) tolower(s) else s},
sep = "", collapse = " " )
sapply(strsplit(s, split = " "), cap, USE.NAMES = !is.null(names(s)))
}

saveXLSX = function(df, file, sheetName = "Sheet1", col.names = TRUE, row.names = TRUE, 
                    append = FALSE, showNA = TRUE, overwrite=TRUE){
  #This function is a simple wrapper around write.xlsx from the xlsx package
  #This version automatically overwrites any existing results
  if(file.exists(file)){
    wb = loadWorkbook(file)
    sheets = getSheets(wb)
    sheetExists = sheetName %in% names(sheets)
    #browser()
    if(sheetExists & overwrite){
      removeSheet(wb, sheetName=sheetName)
      newsheet = createSheet(wb, sheetName=sheetName)
      addDataFrame(x=df, sheet=newsheet, col.names=col.names, row.names=row.names, showNA=showNA)
      saveWorkbook(wb, file)
    } else {
      write.xlsx(as.data.frame(df), file=file, sheetName=sheetName, col.names=col.names, row.names=row.names, showNA=showNA, append=append)
    }
  } else {
    write.xlsx(as.data.frame(df), file=file, sheetName=sheetName, col.names=col.names, row.names=row.names, showNA=showNA, append=append)
  }
}

###############################
#Basic meta-analysis utilities
reDataToDirectMA = function(input.df, dataType){
  #Rearrange data from gemtc input format into a format suitable for direct meta-analysis
  #dataType - Character string indicating what type of data has been provided. 
  # Must be one of: 'treatment difference', 'binary'
  
  #get a list of unique study IDs and count how many
  studyID = unique(input.df$study)
  
  for(s in studyID){
    #pull out the current study
    study = filter(input.df, study==s)
    
    #identify the set of all possible pairwise comparisons in this study
    comparisons = combn(study$treatment, 2)
    
    #rearrange treatment difference data
    if(dataType == 'treatment difference'){
      #set up a temporary data frame
      df = data_frame(StudyName=NA, study=NA, comparator=NA, treatment=NA, diff=NA, std.err=NA,
                      NumberAnalysedComparator=NA, NumberAnalysedTreatment=NA, ComparatorName=NA, TreatmentName=NA)
      
      #loop through the set of comparisons and rearrange the data
      for(i in 1:ncol(comparisons)){
        comp = filter(study, study$treatment %in% comparisons[,i])
        
        #only report the direct comparisons as reported in the data
        if(any(is.na(comp$diff))){
          df[i, 'StudyName'] = as.character(comp$StudyName[1])
          df[i,'study'] = as.integer(s)
          df[i,'treatment'] = as.integer(comp$treatment[2])
          df[i,'comparator'] = as.integer(comp$treatment[1])
          df[i,'diff'] = comp$diff[2]
          df[i,'std.err'] = comp$std.err[2]
          df[i,'NumberAnalysedComparator'] = as.integer(comp$NumberAnalysed[1])
          df[i,'NumberAnalysedTreatment'] = as.integer(comp$NumberAnalysed[2])
          df[i,'ComparatorName'] = as.character(comp$TreatmentName[1])
          df[i,'TreatmentName'] = as.character(comp$TreatmentName[2])
        } 
      }  
      
      if(s == 1){
        directData = df
      } else {
        directData = bind_rows(directData, df)
      }
    }
    
    #rearrange binary data
    if(dataType == 'binary'){
      #set up a temporary data frame
      df = data_frame(StudyName=NA, study=NA, comparator=NA, treatment=NA, NumberEventsComparator=NA,
                      NumberAnalysedComparator=NA, NumberEventsTreatment=NA, NumberAnalysedTreatment=NA,
                      ComparatorName=NA, TreatmentName=NA)
      
      #loop through the set of comparisons and rearrange the data
      for(i in 1:ncol(comparisons)){
        comp = filter(study, study$treatment %in% comparisons[,i])
        df[i, 'StudyName'] = as.character(comp$StudyName[1])
        df[i,'study'] = as.integer(s)
        df[i,'treatment'] = as.integer(comp$treatment[2])
        df[i,'comparator'] = as.integer(comp$treatment[1])
        df[i,'NumberEventsComparator'] = as.integer(comp$responders[1])
        df[i,'NumberAnalysedComparator'] = as.integer(comp$sampleSize[1])
        df[i,'NumberEventsTreatment'] = as.integer(comp$responders[2])
        df[i,'NumberAnalysedTreatment'] = as.integer(comp$sampleSize[2])
        df[i,'ComparatorName'] = as.character(comp$TreatmentName[1])
        df[i,'TreatmentName'] = as.character(comp$TreatmentName[2])
      }  
      
      if(s == 1){
        directData = df
      } else {
        directData = bind_rows(directData, df)
      }
    }
  }
  return(directData)
}

doDirectMeta = function(df, effectCode, dataType, backtransf=FALSE){
  #dataType - Character string indicating what type of data has been provided. 
  # Must be one of: 'treatment difference', 'binary'
  
  #create a list object to store the results
  resList = list()
  
  #identify the set of treatment comparisons present in the data
  comparisons = distinct(df[,3:4])
  
  for(i in 1:nrow(comparisons)){
    #get data for the first comparison
    comp = filter(df, comparator==comparisons$comparator[i], treatment==comparisons$treatment[i])
    
    #run the analysis for different data types
    if(dataType == 'treatment difference'){
      #Generic inverse variance method for treatment differences
      #backtransf=TRUE converts log effect estimates (e.g log OR) back to linear scale
      directRes = metagen(TE=comp$diff, seTE=comp$std.err,sm=effectCode, backtransf=backtransf,
                          studlab=comp$StudyName, n.e=comp$NumberAnalysedTreatment,
                          n.c=comp$NumberAnalysedComparator, label.e=comp$TreatmentName,
                          label.c=comp$ComparatorName)    
    }
    if(dataType == 'binary'){
      #Analysis of binary data provided as n/N
      directRes = metabin(event.e=comp$NumberEventsTreatment, n.e=comp$NumberAnalysedTreatment,
                          event.c=comp$NumberEventsComparator, n.c=comp$NumberAnalysedComparator,
                          sm=effectCode, backtransf=backtransf, studlab=comp$StudyName,
                          label.e=comp$TreatmentName, label.c=comp$ComparatorName,
                          method=ifelse(nrow(comp) >1, "MH", "Inverse"))
    }
    #add the treatment codes to the results object. These will be needed later
    directRes$e.code = comparisons$treatment[i]
    directRes$c.code = comparisons$comparator[i]
    
    #compile the results into a list
    if(length(resList)==0){
      resList[[1]] = directRes
    } else {
      resList[[length(resList)+1]] = directRes
    }
  }
  return(resList)
}

backtransform = function(df){
  #simple function to convert log OR (or HR or RR) back to a linear scale
  #df - a data frame derived from a metagen summary object
  
  #relabel column names to preserve the log results
  colnames(df)[c(1,3:4)] = paste0('log.', colnames(df)[c(1,3:4)])
  colnames(df)[2] = paste0(colnames(df)[2], '.log' )
  #exponentiate the effect estimate and CI
  df$TE = exp(df$log.TE)
  df$lower = exp(df$log.lower)
  df$upper = exp(df$log.upper)
  df
}

extractDirectRes = function(metaRes, effect, intervention='Int', comparator='Con',
                            interventionCode=NA, comparatorCode=NA, backtransf=FALSE){
  #metaRes - an object of class c("metagen", "meta") as returned by the function metagen in the package meta
  #effect - a character string describing the effect estimate, e.g. 'Rate Ratio', 'Odds Ratio', 'Hazard Ratio'
  #intervention - name of the intervention treatment
  #comparator - name of the comparator treatment
  #backtransf - logical indicating whether the results should be exponentiated or not. If the results of the meta-analysis 
  #are log odds ratios set this to TRUE to return the odds ratios. If TRUE this will return both the log estimates and the
  #exponentiated estimates
    
  #create a summary of the results, extract fixed and random
  res = summary(metaRes)
  fixed = as.data.frame(res$fixed)
  random = as.data.frame(res$random)[1:7]
  if(backtransf==TRUE){ #exponentiate the effect estimates if required
    fixed = backtransform(fixed)
    random = backtransform(random)
  }
  
  #if more than one study then this must be a meta-analysis and both fixed and random are expected
  #if there is only one study then use the fixed results as all results are just the result of the 
  #original study
  if(res$k >1){
    df = rbind(fixed, random)
    model=c('Fixed', 'Random')
    df = data.frame('Model'=model, df, stringsAsFactors=FALSE)
  } else {
    df = fixed
    df = data.frame('Model'=NA, df, stringsAsFactors=FALSE)
  }
  studies = paste0(metaRes$studlab, collapse=', ')
  df = data.frame('Intervention'=intervention, 'InterventionCode'=interventionCode,
                  'Comparator'=comparator, 'ComparatorCode'=comparatorCode,
                  'Effect'=effect, df,'Tau.sq'=res$tau, 'method.tau'=res$method.tau,
                  'I.sq'=res$I2, 'n.studies'=res$k, 'studies'=studies,
                  stringsAsFactors=FALSE)
  
}

drawForest = function(meta, showFixed=TRUE, showRandom=TRUE, ...){
  #meta - an object of class c("metagen", "meta") as returned by the function metagen in the package meta
  #showFixed, showRandom - logical indicating whether pooled estimates from fixed or random effects models
  # should be be shown on the forest plot
  
  #work out sensible values for the x-axis limits
  limits = c(meta$lower, meta$upper, meta$lower.fixed, meta$upper.fixed, meta$lower.random, meta$upper.random)
  limits = range(exp(limits))
  xlower = ifelse(limits[1]<0.2, round(limits[1], 1), 0.2)
  xupper = ifelse(limits[2]>5, round(limits[2]), 5)
  xlimits = c(xlower, xupper)
  
  #don't show the pooled estimate if there is only one study
  if(meta$k == 1){
    showFixed = FALSE
    showRandom = FALSE
  }
  
  #forest plot
  forest(meta, hetlab=NULL, text.I2='I-sq', text.tau2='tau-sq', xlim=xlimits, 
         comb.fixed=showFixed, comb.random=showRandom, lty.fixed=0, lty.random=0, ...)
}

buildGraph = function(edgelist, all=FALSE, plotNetwork=FALSE){
  #create a graph object
  #edgelist - a two column matrix describing the connections _from_ the first column _to_ the second column.
  # The values are the treatment numbers
  #all - logical. Describes whether all studies of the same treatments should be shown as separate connections
  # in the graph or collapsed into a single edge. If there are 3 studies comparing the same treatments all=TRUE will
  # show three lines, all=FALSE will only show one (Default)
  #plotNetwork - logical indicating whether to draw a plot of the network or not
  
  #collapse the edges unless explicitly requested not to
  if(all != TRUE){
    el = as.matrix(distinct(edgelist))
  } else {
    el = as.matrix(edgelist)
  }  
  g = graph_from_edgelist(el, directed=TRUE)
  
  if(plotNetwork==TRUE){
    plot(g, layout=igraph::layout.fruchterman.reingold)
  }
  
  return(g)
}

getPathsByLength = function(g, len=2) {
  #Function to return all combinations of 3 connected treatments
  #identifies all paths of length 2 (3 nodes, 2 edges)
  #If there are multiple ways to connect two treatments then all possible combinations are returned
  #g - a graph object as produced by igraph
  #len - the length of path to be identified (Default=2). Connecting 3 nodes has a path length of two
  sp = shortest.paths(g)
  sp[upper.tri(sp,TRUE)] = NA
  wp = which(sp==len, arr.ind=TRUE)
  mapply(function(a,b) get.all.shortest.paths(g, from=a, to=b, mode='all'), wp[,1], wp[,2])
}

bucher = function(abTE, se.abTE, cbTE, se.cbTE, effect, model,
                  intervention, comparator, common, backtransf=FALSE,
                  ab.studies, cb.studies){
  #abTE - treatment effect for a vs b. e.g. log OR, log HR, mean difference
  #se.abTE - standard error of the treatment effect for a vs b, e.g. se of log OR
  #cbTE - treatment effect for c vs b. e.g. log OR, log HR, mean difference
  #se.cbTE - standard error of the treatment effect for c vs b, e.g. se of log OR
  #effect - Character string describing the effect measure, e.g. 'Rate Ratio' or 'log Odds Ratio'
  #model - Character string indicating whether abTE and cbTE come from a fixed effect model or a
  #   random effect model
  #intervention - Character string. Name of the intervention treatment
  #comparator - Character string. Name of the comparator treatment
  #common - Character string. Name of the common comparator, e.g. placebo
  #backtransf - logical indicating whether the results should be exponentiated or not. If abTE and cbTE are 
  #on the log scale set this to TRUE to return the exponentiated results. If TRUE this will return both the
  #log estimates and the exponentiated estimates
  #
  #effect, model, intervention, comparator and common are used for labels only. 
  #The results depend only on abTE, cbTE and the respective standard errors.
  
  acTE = abTE - cbTE #treatment effect
  se.acTE = sqrt(se.abTE^2 + se.cbTE^2) #standard error - sqrt(sum of the two variances)
  log.lower = acTE - (1.96*se.acTE) #calculate confidence intervals
  log.upper = acTE + (1.96*se.acTE)
  df = data.frame('log.TE.ind'=acTE, 'log.lower.ind'=log.lower, 'log.upper.ind'=log.upper, 'se.log.TE.ind'=se.acTE)
  
  #exponentiate if required
  if(backtransf==TRUE){
    eTE = exp(df[,1:3])
    colnames(eTE) = gsub('log.', '', colnames(eTE), fixed=TRUE)
    df = cbind(df, eTE)
  }
  
  #identify and count the studies involved
  ab.studies = strsplit(as.character(ab.studies), split=', ')[[1]]
  cb.studies = strsplit(as.character(cb.studies), split=', ')[[1]]
  studies = c(ab.studies,cb.studies)
  studies = unique(studies)
  n.studies = length(studies)
  studies = paste0(studies, collapse=', ')
  
  #tag on some labels
  df = data.frame('Intervention'=intervention, 'Comparator'=comparator, 'Common'=common,
                  'Effect'=effect, 'Model'=model, df, 'n.studies'=n.studies,'Studies'=studies)
}

doBucher = function(comparisons, direct, effectType='all', backtransf=FALSE){
  #comparisons - a data frame with four columns: StudyName, study, comparator, treatment. Describes the treatment comparisons
  # present in the dataset. study, comparator and treatment must be numbers. For example, study = 4, comparator=1, treatment=2
  # represents the compartison of treatment 2 vs treatment 1 in study 4
  #direct - a data frame containing the results of direct head-to-head meta-analysis for the treatment comparisons of interest
  # if only one study is available for a given comparison then the result of that study should be used. This data frame can be
  # created by using doDirectMeta and extractDirectRes in order
  #effectType - character string indicating what type of results are required. Default is all which will return both fixed effect
  # and random effect results. Alternatives are 'Fixed' or 'Random' (Case sensitive) if only one set of results is required
  
  #take information describing available comparisons from the direct MA dataset
  suppressMessages(require(igraph, quietly=TRUE))
  connections = select(comparisons, from=treatment, to=comparator)
  g = buildGraph(connections)
  
  #get the set of possible BUcher comparisons
  bucherSet = getPathsByLength(g, 2)
  
  for(i in 1:ncol(bucherSet)){
    #for each comparison get the list of possible alternatives
    res = bucherSet[,i]$res
    
    for(j in 1:length(res)){
      #take the current triplet and get the edges connecting those treatments
      #the middle element of the triplet is the common comparator
      triplet = as.numeric(bucherSet[,i]$res[[j]])
      
      #identify the corresponding comparisons in the direct results
      #control which set of results we want. Default is all
      if(effectType == 'all'){
        mod = c('Fixed', 'Random')
      } else {
        mod = effectType
      }
      for(k in 1:length(mod)){
        currComp = filter(direct, (direct$InterventionCode %in% triplet & direct$ComparatorCode %in% triplet), 
                          (Model == mod[k] | is.na(Model)))
        
        #if the common treatment is the intervention rather than the comparator then we need to invert the treatment effect
        ix = currComp$InterventionCode == triplet[2]
        if(any(ix) == TRUE){
          currComp$log.TE[ix] = -currComp$log.TE[ix]
          currComp[ix, c('InterventionCode', 'ComparatorCode')] = currComp[ix, c('ComparatorCode', 'InterventionCode')]
          currComp[ix, c('Intervention', 'Comparator')] = currComp[ix, c('Comparator', 'Intervention')]
        }
        
        #run the comparison
        e = currComp$Effect[1]
        m = currComp$Model[!is.na(currComp$Model)]
        abTE = currComp$log.TE[currComp$InterventionCode==triplet[1]]
        se.abTE = currComp$seTE.log[currComp$InterventionCode==triplet[1]]
        cbTE = currComp$log.TE[currComp$InterventionCode==triplet[3]]
        se.cbTE = currComp$seTE.log[currComp$InterventionCode==triplet[3]]
        int = currComp$Intervention[currComp$InterventionCode==triplet[1]]
        com = currComp$Intervention[currComp$InterventionCode==triplet[3]]
        indMA = bucher(abTE=abTE, se.abTE=se.abTE, 
                       cbTE=cbTE, se.cbTE=se.cbTE,
                       backtransf=backtransf, effect=e, 
                       model=ifelse(length(m)==0, NA, m), 
                       intervention=int, comparator=com,
                       common=currComp$Comparator[1],
                       ab.studies=currComp$studies[1],
                       cb.studies=currComp$studies[2])
        
        #collect FE and RE results
        if(k == 1){
          fere = indMA
        } else {
          fere = rbind(fere, indMA)
        }
        
      }
      #assemble the results
      if(j==1){
        ind.df = fere
      } else {
        ind.df = rbind(ind.df, fere)
      }
    }
    if(i==1){
      df = ind.df
    } else {
      df = rbind(df, ind.df)
    }
  }
  #calculate the inverse comparisons
  df_inv = df
  df_inv[, c('Intervention', 'Comparator')] = df[, c('Comparator', 'Intervention')]
  df_inv$log.TE.ind = -df$log.TE.ind
  df_inv$log.lower.ind = df_inv$log.TE.ind - (1.96*df_inv$se.log.TE.ind)
  df_inv$log.upper.ind = df_inv$log.TE.ind + (1.96*df_inv$se.log.TE.ind)
  df_inv[,10:12] = exp(df_inv[,6:8])
  
  #combine results
  df = rbind(df, df_inv)
  df = arrange(df, desc(Intervention), desc(Comparator))
  
  #fix variable types
  df = factorToCharacter(df)
  
}

###############################
#NMA functions
extractComparison = function(df){
  require(stringr)
  
  #extract the information of which comparison is on each row from the 
  #first column of the winBUGS output and store it in individual columns
  treatments = as.data.frame(str_extract_all(df$node, pattern='\\d+', simplify=TRUE))
  treatments[,1] = as.integer(as.character(treatments[,1]))
  treatments[,2] = as.integer(as.character(treatments[,2]))
  colnames(treatments) = c('tA', 'tB')
  df = tbl_df(bind_cols(treatments, df[,2:ncol(df)]))
}

nameTreatments = function(results, coding, ...){
  #map the treatment names to the numerical codes
  #results = data frame of winbugs output that has already been cleaned up with extract comparisons
  #coding = data frame that maps treatment names to their numerical codes
  
  #drop order column from coding. Makes no sense here
  coding = coding[,1:2]
  
  #match intervention codes to their names
  results = right_join(coding, results, by=c('id'='tB'))
  colnames(results)[1:2] = c('TreatmentB', 'nameB')
  
  #match comparator codes to their names
  results = right_join(coding, results, by=c('id'='tA'))
  colnames(results)[1:2] = c('TreatmentA', 'nameA')
  return(results)
}

makeTab = function(results, coding, rounding=2, reportOrder='default', ...){
  #create summary string
  results[,'effect'] = paste0(round(results$median, rounding), ' (', round(results$CrI_lower,rounding),' to ',round(results$CrI_upper,rounding), ')')
  
  #The order of treatments in the table can be controlled manually
  #using the argument reportOrder.
  #The default is to order the table in the same order the treatments are coded in the network
  if(reportOrder=='custom'){
    coding = arrange(coding, Order)
  }
  #order by treatment coding in the network
  coding$Order = coding$id  
  
  #Create a simple data frame to add output
  reportTab = data.frame('Treatment'=coding$id)    
  
  for(i in 1:nrow(coding)){
    comparator = coding$id[i]
    comp = filter(results, TreatmentA==comparator)
    comp = select(comp, TreatmentB, effect)
    colnames(comp)[ncol(comp)] = coding$id[i]
    reportTab = full_join(reportTab, comp, by=c('Treatment'='TreatmentB'))
  }
  rownames(reportTab) = coding$description
  reportTab = select(reportTab, -Treatment)
  colnames(reportTab) = coding$description
  
  reportTab = as.data.frame(t(reportTab))
}

calcAllPairs = function(mtcRes, expon=FALSE, ...){
  #This function takes two arguments
  #mtcRes is an object of class mtc.result, i.e. the output from mtc.run in the gemtc package
  #expon controls whether the output is exponentiated or not. If the input is log OR (or HR or RR)
  # set this to TRUE to get output on the linear scale
  tid = as.integer(mtcRes$model$network$treatments$id)
  for(t in 1:length(tid)){
    re = suppressWarnings(summary(relative.effect(mtcRes, t1=tid[t], preserve.extra=FALSE)))
    
    stats = re$summaries$statistics
    stats = data.frame(node=rownames(stats), stats, row.names=NULL)
    quan = re$summaries$quantiles
    quan = data.frame(node=rownames(quan), quan[,c(1,3,5)], row.names=NULL)
    colnames(quan) = c('node','CrI_lower', 'median', 'CrI_upper')
    out = full_join(stats, quan, by='node')
    
    if(t==1){
      output = out
    } else {
      output = suppressWarnings(bind_rows(output, out))
    }
  }
  
  if(expon==TRUE){
    output[,c(2,6:8)] = exp(output[,c(2,6:8)])
    output$SD = (output$CrI_upper - output$CrI_lower)/3.92
  } 
  output$Sample = nrow(mtcRes$samples[[1]])
  output = output[,c(1,7,6,8,2:5,9)]
}

extractModelFit = function(mtcRes){
  #mtcRes - an mtc.result object as returned by mtc.run from the gemtc package
  modelSummary = summary(mtcRes)
  dic = data.frame('Mean'=modelSummary$DIC, 'SD'=NA, row.names=c('Dbar', 'pD', 'DIC'))
  if(mtcRes$model$linearModel == 'random'){
    modelSD = modelSummary$summaries$statistics
    ix = rownames(modelSD) == 'sd.d'
    msd = data.frame('Mean'=modelSD[ix,1], 'SD'=modelSD[ix,2], row.names=rownames(modelSD)[ix])
    dic = rbind(msd, dic)
  }
  return(dic)
}

extractResults = function(res, resultsFile, includesPlacebo=FALSE, ...){
  require(xlsx)
  #calculate all pairwise effects
  pairwiseResults = calcAllPairs(res, ...)
  saveXLSX(as.data.frame(pairwiseResults), file=resultsFile, sheetName='Raw',
             showNA=FALSE, row.names=FALSE, append=TRUE)
  
  #extract the comparison info from the first column
  #map treatment names to numbers
  pairwiseResults = extractComparison(pairwiseResults)
  pairwiseResults = nameTreatments(pairwiseResults, ...)
  saveXLSX(as.data.frame(pairwiseResults), file=resultsFile, sheetName='ProcessedAll',
             showNA=FALSE, row.names=FALSE, append=TRUE)
  
  #make a table of all pairwise comparisons and save it as an excel file
  #ALWAYS APPEND=TRUE OR YOU WILL OVERWRITE THE EXISTING RESULTS
  reportTab = makeTab(results=pairwiseResults, ...)
  if(includesPlacebo){
    reportTab = select(reportTab, -matches('Placebo')) #drop the placebo column
  }
  rt = as.data.frame(reportTab, check.names=FALSE)
  saveXLSX(rt, file=resultsFile, sheetName='Report', showNA=FALSE, append=TRUE)
  
  #extract and save the model fit information
  modelFit = extractModelFit(res)
  saveXLSX(modelFit, file=resultsFile, sheetName='DIC', showNA=FALSE, row.names=TRUE, append=TRUE)
  
  return(pairwiseResults)
}

extractTOI = function(df, treatments, toi, intervention=TRUE, orderResults=FALSE){
  #df   - a data frame of results as produced after running extractComparison and nameTreatments
  #toi  - the code number of the treatment of interest (toi) in the network
  #treatments - a table of treatments with columns 'id', 'description', 'Order'
  #The 'Order' column is optional. If present it should be a column of integers describing the order the results
  #should be returned in. Note that the name is case sensitive
  #If you wish to use thise feature set orderResults=TRUE
  #intervention - specifies whether the treatment of interest is an intervention or a comparator (i.e. placebo)
  #orderResults - controls whether the results are returned in a specific order or in the order they were given.
  #If present this should be a vector if integers
  
  #extract the treatment of interest
  if(intervention == TRUE){
    df = filter(df, TreatmentB==toi)
  } else {
    df = filter(df, TreatmentA==toi)
  }
  
  #sort if required
  if(intervention == TRUE && orderResults == TRUE){
    df = left_join(df, treatments[,2:3], by=c('nameA'='description'))
    df = arrange(df, Order)
    df$Order = 1:nrow(df)
  }
  if(intervention == FALSE && orderResults == TRUE){
    df = left_join(df, treatments[,2:3], by=c('nameB'='description'))
    df = arrange(df, Order)
    df$Order = 1:nrow(df)
  }
  
  return(df)
}

saveModelCode = function(mtcRes, modelFile){
  #create some basic descriptive text from the model object
  des = paste0('#Description: ', mtcRes$model$network$description, '\n')
  cat(des, file=modelFile)
  modelType = paste0('#Model Type: ', mtcRes$model$linearModel, ' effects\n')
  cat(modelType, file=modelFile, append=TRUE)
  consistency = paste0('#Consistency Assumption: ', mtcRes$model$type, '\n')
  cat(consistency, file=modelFile, append=TRUE)
  like = paste0('#Likelihood: ', mtcRes$model$likelihood, '\n')
  cat(like, file=modelFile, append=TRUE)
  link = paste0('#Link: ', mtcRes$model$link, '\n')
  cat(link, file=modelFile, append=TRUE)
  nchains = paste0('#Number of chains: ', mtcRes$model$n.chain,'\n')
  cat(nchains, file=modelFile, append=TRUE)
  
  #lastly save the actual model code
  cat(mtcRes$model$code, file=modelFile, append=TRUE)
}

saveDiagnostics = function(mtc, directory){
  #Save some convergence diagnostics as pdf files. Plots are based on ggmcmc package
  #mtc - an object of class mtc.result as returned but mtc.run in the gemtc package
  #directory - a file path indicating the folder to save the results
  
  #if there is no directory to save the plot files then create one
  if(!dir.exists(directory)){dir.create(directory, recursive=TRUE)}
  diagData = ggs(as.mcmc.list(mtc)) #convert data to ggplot friendly format
  f = file.path(directory, 'Traceplot.pdf') #set file name and make traceplots
  ggmcmc(diagData, file=f, plot='traceplot', param_page=3, simplify_traceplot=0.25)
  f = file.path(directory, 'Density.pdf') #set file name and make density plots
  ggmcmc(diagData, file=f, plot='density', param_page=3)
  f = file.path(directory, 'Autocorrelation.pdf') #set file name and make autocorrelation plots
  ggmcmc(diagData, file=f, plot='autocorrelation', param_page=3)
}

plotEstimates = function(df, yvar, xvar='median', lowLimit='CrI_lower', hiLimit='CrI_upper',
                         xlabel='Effect Estimate', noEffectLine=1, yOrder=NA){
  require(ggplot2)
  require(scales)
  require(gridExtra)
  
  df = as.data.frame(df)
  if(is.na(yOrder)){
    df[,yvar] = factor(df[,yvar], levels = sort(unique(df[,yvar])))
    p = ggplot(df) + geom_point(aes_string(x=xvar, y=yvar), size=4)
    p = p + geom_errorbarh(aes_string(x=xvar, y=yvar, xmax = hiLimit, xmin = lowLimit), height = 0.15)
    p = p + scale_y_discrete(limits=rev(levels(df[,yvar])))
  } else {
    p = ggplot(df) + geom_point(aes_string(x=xvar, y=yOrder), size=4)
    p = p + geom_errorbarh(aes_string(x=xvar, y=yOrder, xmax = hiLimit, xmin = lowLimit), height = 0.15)
    p = p + scale_y_reverse(breaks=df[,yOrder], labels=df[,yvar])
  }
  
  #set sensible scale for x-axis
  #default is for ratio measures (OR, HR etc).
  #Mean difference needs special handling
  xrange = range(df[,lowLimit], df[,hiLimit])
  if(xlabel != 'Mean Difference' | xlabel != 'Effect Estimate'){
    xrange[1] = ifelse(xrange[1] > 0, 0, NA)
    xrange[2] = ifelse(xrange[2] < 2, 2, NA)
  } else {
    xrange = c(NA,NA)
  }
  
  p = p + scale_x_continuous(breaks=pretty_breaks(n=10), limits=xrange)
  p = p + labs(x=xlabel)
  p = p + geom_vline(xintercept=noEffectLine)
  p = p + theme_bw()
  p = p + theme(axis.title.y=element_blank(), panel.grid=element_blank(),
                axis.text=element_text(size=12), axis.title=element_text(size=12, vjust=0),
                panel.border=element_rect(colour='black'))
  
}

#######################
#Deprecated
#######################
# prepForWB = function(df, analysisType='TreatmentEffects'){
#   
#   if(analysisType == 'TreatmentEffects'){
#     #check that standard column names have been set
#     columns = c('StudyID', 'Treatment', 'EffectEst', 'SEEffectEst', 'SEBaselineArm', 'Code')
#     if(all(colnames(df)!= columns)){
#       m = paste('Check that column names have been set correctly. The required columns are:',
#                 'StudyID', 'Treatment', 'EffectEst', 'SEEffectEst', 'SEBaselineArm', sep=' ' )
#       message(m)
#     }
#     
#     #work out structure for winbugs input data
#     #find largest number of study arms
#     nArms = count(df, StudyID)
#     ix = order(nArms$StudyID)
#     nArms = nArms[ix,]
#     maxArms = max(nArms$n, na.rm = TRUE)
#     
#     
#     #number of studies
#     studies = sort(unique(df$StudyID))
#     ns = length(studies)
#     
#     #create an object to hold the data in winbugs format
#     nc = (2*maxArms)+ (2*(maxArms-1)) +3
#     wbdata = as_data_frame(as.data.frame(matrix(data=NA, nrow=ns, ncol=nc)))
#     #ugly code to create awful winbugs style column names
#     colnames(wbdata) = c('StudyID', 
#                          paste0('treatment', 1:maxArms),
#                          paste0('t[,', 1:maxArms, ']'), 
#                          paste0('y[,', 2:maxArms, ']'),
#                          paste0('se[,', 2:maxArms, ']'),
#                          'na[]', 'V[]'
#     )
#     
#     #process each study to rearrange the data
#     for(i in 1:length(studies)){
#       study = filter(df, StudyID==studies[i])
#       
#       #number of arms
#       na = nArms$n[i]
#       wbdata[i,'na[]'] = na
#       
#       #treatment labels
#       tr = select(study, StudyID, Treatment, Code)
#       tr = spread(tr, Code, Treatment)
#       wbdata[i, 1:ncol(tr)] = tr[1,]
#       
#       #treatment codes
#       ti = grep('t\\[', colnames(wbdata))[1:na]
#       wbdata[i, ti] = sort(study$Code)
#       
#       #treatment effects
#       y = select(study, StudyID, EffectEst, Code)
#       y = spread(y, Code, EffectEst)
#       yi = grep('y\\[', colnames(wbdata))[1:(na-1)]
#       wbdata[i,yi] = y[,(3:(3+(na-2)))]
#       
#       #standard error of treatment effects
#       se = select(study, StudyID, SEEffectEst, Code)
#       se = spread(se, Code, SEEffectEst)
#       sei = grep('se\\[', colnames(wbdata))[1:(na-1)]
#       wbdata[i,sei] = se[,(3:(3+(na-2)))]
#       
#       #for multi-arm trials we also need the variance in the baseline arm
#       #tests whether the multiarm argument is true and a column has been set for
#       #the standard error of the baseline arm
#       if(na>2){
#         control = which(is.na(study$EffectEst))
#         seb = study[control, 'SEBaselineArm']
#         vb = seb^2
#         wbdata[i, 'V[]'] = vb
#       } 
#     }
#   }
#   
#   #processing for binary data
#   if(analysisType=='Binary'){
#     #check that standard column names have been set
#     columns = c('StudyID', 'Treatment', 'Events', 'TotalN', 'Code')
#     if(all(colnames(df)!= columns)){
#       m = paste('Check that column names have been set correctly. The required columns are:',
#                 'StudyID', 'Treatment', 'Events', 'TotalN', 'Code', sep=' ' )
#       message(m)
#     }
#     
#     #work out structure for winbugs input data
#     #find largest number of study arms
#     nArms = count(df, StudyID)
#     ix = order(nArms$StudyID)
#     nArms = nArms[ix,]
#     maxArms = max(nArms$n, na.rm = TRUE)
#     
#     #number of studies
#     studies = sort(unique(df$StudyID))
#     ns = length(studies)
#     
#     #create an object to hold the data in winbugs format
#     nc = 1 + (2*maxArms) + (2*maxArms) + 1
#     wbdata = as_data_frame(as.data.frame(matrix(data=NA, nrow=ns, ncol=nc)))
#     #ugly code to create awful winbugs style column names
#     r = paste0('r[,', 1:maxArms, ']')
#     n = paste0('n[,', 1:maxArms, ']')
#     rn = c(rbind(r, n))
#     colnames(wbdata) = c('StudyID', 
#                          paste0('treatment', 1:maxArms),                          
#                          rn,
#                          paste0('t[,', 1:maxArms, ']'),
#                          'na[]')
#     
#     for(i in 1:length(studies)){
#       study = filter(df, StudyID==studies[i])
#       
#       #number of arms
#       na = nArms$n[i]
#       wbdata[i,'na[]'] = na
#       
#       #treatment labels
#       tr = select(study, StudyID, Treatment, Code)
#       tr = spread(tr, Code, Treatment)
#       wbdata[i, 1:ncol(tr)] = tr[1,]
#       
#       #treatment codes
#       ti = grep('t\\[', colnames(wbdata))[1:na]
#       wbdata[i, ti] = sort(study$Code)
#       
#       #events per arm
#       r = select(study, StudyID, Events, Code)
#       r = spread(r, Code, Events)
#       ri = grep('r\\[', colnames(wbdata))[1:na]
#       wbdata[i,ri] = r[1,2:ncol(r)]
#       
#       #Total N per arm
#       n = select(study, StudyID, TotalN, Code)
#       n = spread(n, Code, TotalN)
#       ni = grep('n\\[', colnames(wbdata))[1:na]
#       wbdata[i,ni] = n[1,2:ncol(n)]
#       
#     }
#   }
#   #order the rows so that multiarm studies are at the end
#   wbdata = arrange(wbdata, `na[]`, StudyID)
# }
#############################################