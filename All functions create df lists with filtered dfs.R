library(readxl)
library(writexl)
library(tidyverse)
library(dplyr)
library(stringr)
library(janitor)


#read the kNook excel
read_knook_requests <- function(xlfilepath) {
  #'Import Knook excel file into R, coerce to a Data Frame
  #'assigns the correct type to each column of the Knook file,
  #'and cleans names of all columns.
  #'returns KnookReqdf, a Data Frame of the Knook Requests
  
  #'Needs xlfilepath as argument.
  #'xlfilepath is the path to the .xlsx Knook file,
  #'usually something like: /Users/blancarincondearellano/Downloads/NDC_Partnership_Data-2022-11-12_2045.xlsx
  
  rawKnookReq <- read_excel(xlfilepath)
  KnookReqdf <- clean_names(rawKnookReq)
  KnookReqdf <- as.data.frame(KnookReqdf)
  KnookReqdf$estimated_cost <- as.numeric(KnookReqdf$estimated_cost)
  KnookReqdf$date_partnership_plan_validated_rsl_received <- as.Date(KnookReqdf$date_partnership_plan_validated_rsl_received, format = "%m/%d/%Y")
  KnookReqdf$start_date <- as.Date(KnookReqdf$start_date, format = "%m/%d/%Y")
  KnookReqdf$end_date <- as.Date(KnookReqdf$end_date, format = "%m/%d/%Y")
  KnookReqdf$date_last_updated <- as.Date(KnookReqdf$date_last_updated, format = "%m/%d/%Y")
  
  #seems to be giving an error wehere all countries appear in each row
  #KnookReqdf$country <- toString(KnookReqdf$country)
  
  #remove parenthesis from determined columns (not text heavy columns like outcome, output, etc.)
  #first parenthesis
  KnookReqdf$country <- str_replace_all(KnookReqdf$country, "[//(]" ,"")
  KnookReqdf$key_topics_covered <- str_replace_all(KnookReqdf$key_topics_covered, "[//(]" ,"")
  KnookReqdf$activity_types <- str_replace_all(KnookReqdf$activity_types, "[//(]" ,"")
  KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support <- str_replace_all(KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support,"[//(]" ,"")
  KnookReqdf$ip_d_ps_that_have_indicated_indicative_support <- str_replace_all(KnookReqdf$ip_d_ps_that_have_indicated_indicative_support, "[//(]" ,"")
  KnookReqdf$lead_domestic_agency <- str_replace_all(KnookReqdf$lead_domestic_agency, "[//(]" ,"")
  #second parenthesis
  KnookReqdf$country <- str_replace_all(KnookReqdf$country, "[//)]" ,"")
  KnookReqdf$key_topics_covered <- str_replace_all(KnookReqdf$key_topics_covered, "[//)]" ,"")
  KnookReqdf$activity_types <- str_replace_all(KnookReqdf$activity_types, "[//)]" ,"")
  KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support <- str_replace_all(KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support,"[//)]", "")
  KnookReqdf$ip_d_ps_that_have_indicated_indicative_support <- str_replace_all(KnookReqdf$ip_d_ps_that_have_indicated_indicative_support, "[//)]", "")
  KnookReqdf$lead_domestic_agency <- str_replace_all(KnookReqdf$lead_domestic_agency, "[//)]" ,"")
  
  #also + signs because ugh
  KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support <- str_replace_all(KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support,"[+]", "")
  KnookReqdf$ip_d_ps_that_have_indicated_indicative_support <- str_replace_all(KnookReqdf$ip_d_ps_that_have_indicated_indicative_support, "[+]", "")
  
  return(KnookReqdf)
}


#filter data frames by variable (lists of df) and print summary analysis
request_source_df <- function(KnookReqdf){
  request_source_df_list <- list()
  request_source_list <- unique(KnookReqdf$request_source)
  
  for (s in request_source_list) {
    filtered_df <- KnookReqdf %>%
      filter(str_detect(request_source, s))
    request_source_df_list[[s]] <- filtered_df
    cat("There are", nrow(filtered_df), "requests under", s, "\n")
  }
  names(request_source_df_list) <- request_source_list
  return(request_source_df_list)
}
sectors_df <- function(KnookReqdf){
  sectors_list_not_split <- list(KnookReqdf$sectors)
  sectors_list_not_split <- sectors_list_not_split[[1]]
  sectors_list_not_split <- sectors_list_not_split[!is.na(sectors_list_not_split)]
  sectors_list <- list()
  sectors_list <- sort(unique(unlist(strsplit(sectors_list_not_split, ", "))))
  sectors_df_list <- list()
  
  for (s in sectors_list) {
    filtered_df <- KnookReqdf %>%
      filter(str_detect(sectors, s))
    sectors_df_list[[s]] <- filtered_df
    cat("There are", nrow(filtered_df), "requests under", s, "\n")
  }
  names(sectors_df_list) <- sectors_list
  return(sectors_df_list)
}
activity_types_analysis_df <- function(KnookReqdf){
  activity_types_list_not_split <- list(KnookReqdf$activity_type)
  activity_types_list_not_split <- activity_types_list_not_split[[1]]
  activity_types_list_not_split <- activity_types_list_not_split[!is.na(activity_types_list_not_split)]
  activity_types_list <- list()
  activity_types_list <- sort(unique(unlist(strsplit(activity_types_list_not_split, ", "))))
  activities_df_list <- list()
  
  for (activity in activity_types_list){
    filtered_df <- KnookReqdf %>%
      filter(str_detect(activity_types, activity))
    activities_df_list[[activity]] <- filtered_df
    cat("There are", nrow(filtered_df), "requests under", activity, "\n")
  }
  names(activities_df_list) <- activity_types_list
  return(activities_df_list)
}
key_topics_analysis_df <- function(KnookReqdf){
  key_topics_list_not_split <- list(KnookReqdf$key_topics)
  key_topics_list_not_split <- key_topics_list_not_split[[1]]
  key_topics_list_not_split <- key_topics_list_not_split[!is.na(key_topics_list_not_split)]
  key_topics_list <- list()
  key_topics_list <- sort(unique(unlist(strsplit(key_topics_list_not_split, ", "))))
  topics_df_list <- list()
  
  for (topic in key_topics_list){
    filtered_df <- KnookReqdf %>%
      filter(str_detect(key_topics_covered, topic))
    topics_df_list[[topic]] <- filtered_df
    cat("There are", nrow(filtered_df), "requests under", topic, "\n")
  }
  names(topics_df_list) <- key_topics_list
  return(topics_df_list)
}
country_analysis_df <- function(KnookReqdf){
  countries_df_list <- list()
  countries_list <- unique(KnookReqdf$country)
  cat("Number of countries that requested financial support: ", length(countries_list), "\n")
  
  for (c in countries_list) {
    filtered_df <- KnookReqdf %>%
      filter(str_detect(country, c))
    countries_df_list[[c]] <- filtered_df
    cat("There are", nrow(filtered_df), "requests from", c, "\n")
  }
  names(countries_df_list) <- countries_list
  return(countries_df_list)
}
regional_analysis_df <- function(KnookReqdf){
  #prints table of requests by region
  regions_list <- list()
  regions_list <- unique(KnookReqdf$region)
  regions_df_list <- list()
  
  for (reg in regions_list) {
    filtered_df <- KnookReqdf %>%
      filter(str_detect(region, reg))
    regions_df_list[[reg]] <- filtered_df
  }
  names(regions_df_list) <- regions_list
  
  #printing the summary table about regional distribution
  cat("Regional distribution of submitted requests:")
  regions_unsorted <- table(KnookReqdf$region)
  regions_table <- regions_unsorted[order(regions_unsorted, decreasing = TRUE)]
  print(regions_table)
  
  return(regions_df_list)
}
support_status_analysis_df <- function(KnookReqdf){
  support_status_names_lists <- list("supported", "not_supported", "partially_supported", "indicative", "not_consolidated")
  support_status_df_list <- list()
  
  total_requests <- nrow(KnookReqdf)
  
  supported <- KnookReqdf %>% filter(str_detect(is_this_request_supported, "Yes"))
  cat("Number of requests supported:", nrow(supported), "\n")
  cat(nrow(supported)/total_requests*100, "% of all requests \n")
  not_supported <- KnookReqdf %>% filter(str_detect(is_this_request_supported, "No"))
  cat("Number of requests still unsupported:", nrow(not_supported), "\n")
  cat(nrow(not_supported)/total_requests*100, "% of all requests \n")
  partially_supported <- KnookReqdf %>% filter(str_detect(is_this_request_supported, "Partial"))
  cat("Number of requests partially supported:", nrow(partially_supported), "\n")
  cat(nrow(partially_supported)/total_requests*100, "% of all requests \n")
  indicative <- KnookReqdf %>% filter(str_detect(is_this_request_supported, "Indicative"))
  cat("Number of requests with indicative support:", nrow(indicative), "\n")
  cat(nrow(indicative)/total_requests*100, "% of all requests \n")
  not_consolidated <- KnookReqdf %>% filter(str_detect(is_this_request_supported, "Partner responses not yet consolidated and reflected"))
  cat("Number of requests with unconsolidated support:", nrow(not_consolidated), "\n")
  cat(nrow(not_consolidated)/total_requests*100, "% of all requests \n")
  
  support_status_df_list[[1]] <- supported
  support_status_df_list[[2]] <- not_supported
  support_status_df_list[[3]] <- partially_supported
  support_status_df_list[[4]] <- indicative
  support_status_df_list[[5]] <- not_consolidated
  
  names(support_status_df_list) <- support_status_names_lists
  
  return(support_status_df_list)
}
focus_area_analysis_df <- function(KnookReqdf){
  focus_area_names_lists <- list("Cross-Cutting", "Mitigation", "Adaptation")
  focus_area_df_list <- list()
  
  total_requests <- nrow(KnookReqdf)
  
  cross_cutting <- KnookReqdf %>% filter(str_detect(focus_area, "Cross-Cutting"))
  cat("Number of cross-cutting:", nrow(cross_cutting), "\n")
  cat(nrow(cross_cutting)/total_requests*100, "% of all requests \n")
  mitigation <- KnookReqdf %>% filter(str_detect(focus_area, "Mitigation"))
  cat("Number of mitigation requests:", nrow(mitigation), "\n")
  cat(nrow(mitigation)/total_requests*100, "% of all requests \n")
  adaptation <- KnookReqdf %>% filter(str_detect(focus_area, "Adaptation"))
  cat("Number of adaptation requests:", nrow(adaptation), "\n")
  cat(nrow(adaptation)/total_requests*100, "% of all requests \n")
  
  focus_area_df_list[[1]] <- cross_cutting
  focus_area_df_list[[2]] <- mitigation
  focus_area_df_list[[3]] <- adaptation
  
  names(focus_area_df_list) <- focus_area_names_lists
  
  return(focus_area_df_list)
}
lead_agency_MOF_df <- function(KnookReqdf){
  #agency_list_not_split <- list(KnookReqdf$lead_domestic_agency)
  #agency_list_not_split <- agency_list_not_split[[1]]
  #agency_list_not_split <- agency_list_not_split[!is.na(agency_list_not_split)]
  #agency_names_list <- list()
  #agency_names_list <- sort(unique(unlist(strsplit(agency_list_not_split, ", "))))
  #lead_agency_df_list <- list()
  
  agency_names_list <- list("Ministry of Finance",
                            "Min Finanças",
                            "Ministry of finance",
                            "Ministère des finances",
                            "Ministry of the Economy and Finance",
                            "Ministry of Economy and Finance",
                            "MoF",
                            "MoEF",
                            "MEF")
  lead_agency_df_list <- list()
  
  total_requests <- nrow(KnookReqdf)
  
  for (partner in agency_names_list) {
    filtered_df <- KnookReqdf %>%
      filter(str_detect(lead_domestic_agency, partner))
    lead_agency_df_list[[partner]] <- filtered_df
    cat("Number of requests from", partner,":", nrow(lead_agency_df_list[[partner]]), "\n")
    cat(nrow(filtered_df)/total_requests*100, "% of all requests \n")
  }
  #names(lead_agency_df_list) <- agency_names_list
  MOF_agency_one_df <- do.call("rbind", lead_agency_df_list)
  
  return(MOF_agency_one_df)
}
partner_support_confirmed_analysis_df <- function(KnookReqdf){
  partners_list_not_split <- list(KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support)
  partners_list_not_split <- partners_list_not_split[[1]]
  partners_list_not_split <- partners_list_not_split[!is.na(partners_list_not_split)]
  partner_names_list <- list()
  partner_names_list <- sort(unique(unlist(strsplit(partners_list_not_split, ", "))))
  partner_support_df_list <- list()
  
  total_requests <- nrow(KnookReqdf)
  
  for (partner in partner_names_list) {
    filtered_df <- KnookReqdf %>%
      filter(str_detect(ip_d_ps_that_have_indicated_confirmed_support, partner))
    partner_support_df_list[[partner]] <- filtered_df
    cat("Number of requests supported (confirmed) from", partner,":", nrow(partner_support_df_list[[partner]]), "\n")
    cat(nrow(filtered_df)/total_requests*100, "% of all requests \n")
  }
  names(partner_support_df_list) <- partner_names_list
  
  return(partner_support_df_list)
}
partner_support_confirmed_analysis_table <- function(KnookReqdf){
  partners_list_not_split <- list(KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support)
  partners_list_not_split <- partners_list_not_split[[1]]
  partners_list_not_split <- partners_list_not_split[!is.na(partners_list_not_split)]
  partner_names_list <- list()
  partner_names_list <- sort(unique(unlist(strsplit(partners_list_not_split, ", "))))
  partner_support_df_list <- list()
  
  partner_num_req_supported_list <- c()
  partner_percent_req_supported_list <- c()
  partner_support_df_list <- list()
  
  total_requests <- nrow(KnookReqdf)
  
  for (partner in partner_names_list) {
    filtered_df <- KnookReqdf %>%
      filter(str_detect(ip_d_ps_that_have_indicated_confirmed_support, partner))
    partner_support_df_list[[partner]] <- filtered_df
    partner_num_req_supported_list[partner] <- nrow(partner_support_df_list[[partner]])
    partner_percent_req_supported_list[partner] <-nrow(filtered_df)/total_requests*100
  }
  names(partner_support_df_list) <- partner_names_list
  
  partner_support_table <- data_frame(partner_names_list, partner_num_req_supported_list, partner_percent_req_supported_list)
  return(partner_support_table)
}
partner_support_indicative_analysis_df <- function(KnookReqdf){
  partners_list_not_split <- list(KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support)
  partners_list_not_split <- partners_list_not_split[[1]]
  partners_list_not_split <- partners_list_not_split[!is.na(partners_list_not_split)]
  partner_names_list <- list()
  partner_names_list <- sort(unique(unlist(strsplit(partners_list_not_split, ", "))))
  partner_support_df_list <- list()
  
  total_requests <- nrow(KnookReqdf)
  
  for (partner in partner_names_list) {
    filtered_df <- KnookReqdf %>%
      filter(str_detect(ip_d_ps_that_have_indicated_indicative_support, partner))
    partner_support_df_list[[partner]] <- filtered_df
    cat("Number of requests supported (confirmed) from", partner,":", nrow(partner_support_df_list[[partner]]), "\n")
    cat(nrow(filtered_df)/total_requests*100, "% of all requests \n")
  }
  names(partner_support_df_list) <- partner_names_list
  
  return(partner_support_df_list)
}
partner_support_indicative_analysis_table <- function(KnookReqdf){
  partners_list_not_split <- list(KnookReqdf$ip_d_ps_that_have_indicated_confirmed_support)
  partners_list_not_split <- partners_list_not_split[[1]]
  partners_list_not_split <- partners_list_not_split[!is.na(partners_list_not_split)]
  partner_names_list <- list()
  partner_names_list <- sort(unique(unlist(strsplit(partners_list_not_split, ", "))))
  partner_support_df_list <- list()
  
  total_requests <- nrow(KnookReqdf)
  
  for (partner in partner_names_list) {
    filtered_df <- KnookReqdf %>%
      filter(str_detect(ip_d_ps_that_have_indicated_indicative_support, partner))
    partner_support_df_list[[partner]] <- filtered_df
    partner_num_req_supported_list[partner] <- nrow(partner_support_df_list[[partner]])
    partner_percent_req_supported_list[partner] <-nrow(filtered_df)/total_requests*100
  }
  names(partner_support_df_list) <- partner_names_list
  
  partner_support_table <- data_frame(partner_names_list, partner_num_req_supported_list, partner_percent_req_supported_list)
  return(partner_support_table)
}

filter_by_date_df <- function(KnookReqdf, date, after = TRUE){
  if (after == TRUE){
    new_df_by_date <- KnookReqdf %>%
      filter(date_partnership_plan_validated_rsl_received >= as.Date(date, format = "%m/%d/%Y"))
  } else {
    new_df_by_date <- KnookReqdf %>%
      filter(date_partnership_plan_validated_rsl_received <= as.Date(date, format = "%m/%d/%Y"))
  }
  return(new_df_by_date)
}

#formula for manual analysis by Finance category (new column created)
categories_analysis_df <- function(categorisedKnookReqdf){
  categories_df_list <- list()
  categories_list <- unique(categorisedKnookReqdf$category)
  
  for (c in categories_list) {
    filtered_df <- categorisedKnookReqdf %>%
      filter(str_detect(category, c))
    categories_df_list[[c]] <- filtered_df
    cat("There are", nrow(filtered_df), "requests under", c, "\n")
  }
  names(categories_df_list) <- categories_list
  return(categories_df_list)
}
subcategories_analysis_df <- function(categorisedKnookReqdf){
  subcategories_df_list <- list()
  subcategories_list <- unique(categorisedKnookReqdf$subcategory_what)
  
  for (s in subcategories_list) {
    filtered_df <- categorisedKnookReqdf %>%
      filter(str_detect(subcategory_what, s))
    subcategories_df_list[[s]] <- filtered_df
    cat("There are", nrow(filtered_df), "requests under", s, "\n")
  }
  names(subcategories_df_list) <- subcategories_list
  return(subcategories_df_list)
}

#Create an xfilepath for the relevant excel file to be used
xlfilepath_Aug_2023 <- "/Users/blancarincondearellano/Downloads/NDC_Partnership_Data-2023-08-16_1752.xlsx"
AllFinReqdf <- read_knook_requests(xlfilepath_Aug_2023)
